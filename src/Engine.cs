// 候选引擎：攒池子、匹配、排序、决定弹不弹、以及"选完之后该补什么"。
//
// ★这个类不引用 WinForms 的任何东西 —— 窗口通过回调接进来。
//   于是引擎的全部行为（弹不弹、排第几、补括号还是补空格）都能在没有界面的
//   情况下测。
//
// ★v132：排序与匹配彻底对齐 Python 版（那一版是用户在真机上一条条调出来的）。
//   最关键的修正是【排序键的顺序】：
//     旧：作用域 > 命中位置 > 跨度 > 最近使用 > ...
//     新：命中质量 > 命中位置 > 跨度 > 作用域 > 最近使用 > 内建 >
//         作用域优先 > 名字长度
//   旧顺序把"工程内的名字"一律顶到最前，于是打了 `msgbox` 也会因为某个工程内
//   的 8 字母变量只是"松散命中"而排在 MsgBox 前面 —— 用户看到的就是
//   "最像的那个不在第一个"。Python 版的注释里明确写过"作用域放最前 = 会顶掉
//   匹配得更好的候选，明确不要"。

using System;
using System.Collections.Generic;
using System.Linq;

namespace VbeComplete
{
    internal sealed class Candidate
    {
        public string Name;
        public string Kind;        // func / prop / var / type / kw —— 图标档位
        public int Pos;            // 首个命中下标（越小越好）
        public int Span;           // 命中跨度（越小越好）
        public int[] Hits;         // 逐字符命中的下标（UI 高亮用；空 = 全都算命中）
        public int HitKind;        // 4 完全相等 / 3 前缀 / 2 词首缩略 / 1 连续块 / 0 分散
        public int Run;            // 最长连续命中段
        public int ScopeKind;      // 4 本过程局部 / 3 本模块 / 2 别的模块 Public
        public int RecentRank;     // 最近使用次数
        public bool IsBuiltin;
    }

    internal sealed class Engine
    {
        public IVbeSource Backend;
        public Recent Recent;
        public WordBank Bank;

        // ★一屏最多几行：与 Python 版 MAX_VISIBLE_ROWS 对齐（15）。候选不足时
        //   窗口自己变矮 —— 用户口径是"显示几行就几行高"，不是固定 15 行留白。
        public readonly int ViewRows = Settings.Int("VIEW_ROWS", 15, 3, 30);
        public readonly bool HideOnComplete = Settings.Flag("HIDE_ON_COMPLETE", true);
        public readonly bool MemberPopup = Settings.Flag("MEMBER_POPUP", false);
        public readonly bool DeclPopup = Settings.Flag("DECL_POPUP", false);

        // UI 回调
        public Action<IReadOnlyList<Candidate>, int, int> OnShow;   // items, selected, top
        public Action OnHide;
        public Action OnMove;                                        // 选中项变了，重画
        /// <summary>
        /// "现在该不该让位给 VBE 自己的成员列表"。非 null 且返回 true 时本次
        /// 触发直接收起 —— 两个列表叠在一起，用户分不清在看哪一个。
        /// 抽成回调是为了让引擎本身不依赖 Win32（自检里传 null 就行）。
        /// </summary>
        public Func<bool> YieldToVbe;

        private List<Candidate> _items = new List<Candidate>();
        private int _sel;
        private int _top;
        public bool Visible { get; private set; }

        // 池：全工程的声明（每个模块一份，按行数缓存）
        private readonly Dictionary<string, ModuleScope> _scopeCache =
            new Dictionary<string, ModuleScope>(StringComparer.OrdinalIgnoreCase);
        private List<ModuleScope> _scopes = new List<ModuleScope>();
        private List<string> _poolNames = new List<string>();
        // 池里哪些名字【必须】通过作用域判定才能出现（= 声明出来的名字）。
        // 内建 / 组件名 / 词库 / "引用过但没声明"这些是结构性的，不受作用域约束。
        private readonly HashSet<string> _poolScoped = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> _procKinds =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private HashSet<string> _typeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string _poolKey;
        private string _poolModule;

        // 刚按 Tab 确认过的那个词。同一行、同一个词不再提示自己 ——
        // Python 版 v90 的那条口径（"按完 Tab 意味着本次录入完成，不要再匹配提示"）。
        private string _acceptedWord;
        private string _acceptedModule;
        private int _acceptedLine;
        private long _acceptedTicks;
        private const long AcceptQuietMs = 2000;

        // 这一次触发时"光标前的词是不是已经打全了"（空格确认要用）
        private bool _wordComplete;

        public IReadOnlyList<Candidate> Items => _items;
        public int SelectedIndex => _sel;
        public int TopIndex => _top;
        public bool WordComplete => _wordComplete;

        // ------------------------------------------------------------------
        // 池子
        // ------------------------------------------------------------------

        /// <summary>
        /// 重建池子。key 带"本模块文本长度 + 全工程签名"：
        /// 本模块一边打字一边重解析（保证自己新写的变量立刻能提示），
        /// 别的模块只在行数变了之后才重解析（不然每次按键都要重扫整个工程）。
        /// </summary>
        private void EnsurePool(string moduleName, string moduleText)
        {
            int sig = 0;
            List<ModuleScope> others = null;
            if (Backend != null)
            {
                others = Backend.ProjectScopes(out sig);
            }

            string key = (moduleName ?? "") + "|" + (moduleText == null ? -1 : moduleText.Length) + "|" + sig;
            if (_poolNames.Count > 0 && _poolKey == key) return;

            var scopes = new List<ModuleScope>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _procKinds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _typeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1) 当前模块：用传进来的文本（最新鲜），不吃缓存
            ModuleScope cur = null;
            if (moduleText != null)
            {
                cur = Scope.Analyze(moduleName ?? "", moduleText);
                _scopeCache[moduleName ?? ""] = cur;
                scopes.Add(cur);
                seen.Add(moduleName ?? "");
            }

            // 2) 其他模块
            if (others != null)
                foreach (var ms in others)
                {
                    if (!seen.Add(ms.Module)) continue;
                    scopes.Add(ms);
                }
            _scopes = scopes;

            // 3) 合并名字
            var names = new List<string>();
            var nameSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _poolScoped.Clear();
            foreach (var ms in scopes)
            {
                foreach (var d in ms.Decls)
                {
                    if (nameSeen.Add(d.Name)) names.Add(d.Name);
                    // ★声明出来的名字一律要过作用域判定：查不到档位 = 当前位置
                    //   看不见它（别的过程的局部 / 别的模块的 Private），必须剔除。
                    _poolScoped.Add(d.Name);
                    if (d.Proc == null && d.KindHint != null && d.KindHint != "type" && d.KindHint != "enum")
                        _procKinds[d.Name] = d.KindHint;
                    if (d.KindHint == "type" || d.KindHint == "enum") _typeNames.Add(d.Name);
                }
                foreach (var t in ms.TypeNames) if (nameSeen.Add(t)) names.Add(t);
                foreach (var f in ms.FreeIdents) if (nameSeen.Add(f)) names.Add(f);
                foreach (var p in ms.Procs)
                {
                    if (nameSeen.Add(p.Name)) names.Add(p.Name);
                    _poolScoped.Add(p.Name);
                    _procKinds[p.Name] = p.Kind;
                }
            }

            // 4) 组件名（模块 / 类模块 / 窗体）
            if (Backend != null)
                foreach (var n in Backend.ComponentNames())
                    if (nameSeen.Add(n)) names.Add(n);

            // 5) 词库里的额外词
            if (Bank != null)
                foreach (var n in Bank.ExtraNames())
                    if (nameSeen.Add(n)) names.Add(n);

            _poolNames = names;
            _poolModule = moduleName ?? "";
            _poolKey = key;
        }

        /// <summary>把每个名字折算成"作用域档位"，key 为小写名。</summary>
        private Dictionary<string, int> BuildRankMap(string curModule, string curProc)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var ms in _scopes)
            {
                bool isCur = string.Equals(ms.Module, curModule, StringComparison.OrdinalIgnoreCase);
                var src = isCur ? ms.Visible(curProc) : ms.Exported();
                foreach (var kv in src)
                {
                    int r = kv.Value.Value;
                    int old;
                    if (map.TryGetValue(kv.Key, out old) && old >= r) continue;
                    map[kv.Key] = r;
                }
            }
            return map;
        }

        /// <summary>把池子按给定源码算出来（自检用；运行期由 Trigger 调）。</summary>
        public void PrimePool(string moduleText, string moduleName)
        {
            EnsurePool(moduleName ?? "", moduleText);
        }

        public void InvalidatePool()
        {
            _poolKey = null;
            _poolNames.Clear();
        }

        // ------------------------------------------------------------------
        // 触发
        // ------------------------------------------------------------------

        /// <summary>
        /// 判断现在该不该弹。manual=true 是 Ctrl+Space（豁免"打全即收"等收束规则）。
        /// 返回 true 表示窗口已显示。
        /// </summary>
        public bool Trigger(bool manual, Snapshot snap = null)
        {
            snap = snap ?? Backend?.Snapshot();
            if (snap == null) { Hide(); return false; }

            // VBE 自己的成员列表正画着 —— 一律让位。让位不是"不提示"：
            // 它的列表比我们的更准（有真类型信息），等它关掉之后下一次触发
            // 我们再补位（Python 版同一套口径）。
            if (!manual && YieldToVbe != null && YieldToVbe()) { Hide(); return false; }

            string line = snap.Text ?? "";
            // ★GetSelection 给的是**显示列**，后面所有取词/判据都按字符列算
            //   —— 含中文/全角的行不换算就会取到空词，连弹窗都不出现。
            int caret = Rules.CharColFromDisp(line, snap.CaretCol);

            if (!manual && Rules.InCommentOrString(line, caret)) { Hide(); return false; }
            if (!manual && !DeclPopup && Rules.InDeclarationPosition(line, caret)) { Hide(); return false; }

            string text = Backend?.ModuleText();
            EnsurePool(snap.Module, text);

            // ---- `对象.` 位置 ----
            // VBE 自己有成员列表（带真类型信息），默认整个让给它。只有在
            // 【我们能确定这个变量的类型】时才补一份 —— 类型拿不到就还是不弹，
            // 绝不拿"全工程候选"去冒充成员列表（那是噪音，比不弹更糟）。
            // `MEMBER_POPUP=1` 是另一条路：不做类型收窄，直接出全工程候选
            // （老行为，留给"就是要看全部名字"的人）。
            List<string> memberPool = null;
            bool atMemberPos = false;
            if (!manual)
            {
                string root;
                int dotCol;
                if (Rules.MemberRoot(line, caret, out root, out dotCol))
                {
                    atMemberPos = true;
                    if (!MemberPopup)
                    {
                        memberPool = MemberPool(root, snap);
                        if (memberPool == null) { Hide(); return false; }
                    }
                }
            }

            int wordStart;
            string word = Rules.WordBefore(line, caret, out wordStart);
            if (word.Length == 0)
            {
                if (memberPool != null || ((manual || MemberPopup) && atMemberPos))
                {
                    // 刚打点号 / 手动唤出：列出全部（0 个字符的 query 匹配一切）
                }
                else if (!manual) { Hide(); return false; }
                else wordStart = caret;      // Ctrl+Space 在空位置：列出全部
            }

            // 刚确认过的词不再提示自己（按 Tab 之后光标处那个词）
            if (!manual && QuietFromAccepted(word, snap)) { Hide(); return false; }

            var ranks = BuildRankMap(snap.Module, ProcOf(snap));

            var hits = new List<Candidate>(64);
            foreach (var name in memberPool ?? _poolNames)
            {
                if (string.IsNullOrEmpty(name)) continue;
                FuzzyHit fh;
                if (!Fuzzy.Match(name, word, out fh)) continue;

                // 单字母只认前缀/精确 —— 否则打一个 `e` 会命中几百个
                if (word.Length == 1 && fh.Kind < 3) continue;

                bool isBuiltin = VbaBuiltins.IsBuiltinFunction(name) ||
                                 VbaBuiltins.IsBuiltinType(name) ||
                                 VbaBuiltins.IsKeyword(name);

                // 2~3 个字符的短输入，内建词汇只接受"首字母相同"的跳步命中：
                //   ve -> VarType 可以；xl -> Explicit 不行（纯噪音）
                if (isBuiltin && fh.Kind == 0 && word.Length >= 2 && word.Length <= 3 &&
                    !string.Equals(name.Substring(0, 1), word.Substring(0, 1),
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                int rank;
                if (memberPool != null)
                {
                    rank = 3;                     // 成员：天然属于当前对象
                }
                else if (!ranks.TryGetValue(name, out rank))
                {
                    // 声明出来的名字查不到档位 = 当前位置根本看不见它
                    // （别的过程的局部变量 / 别的模块的 Private）—— 直接剔除。
                    if (_poolScoped.Contains(name)) continue;
                    rank = 2;
                }

                var c = new Candidate
                {
                    Name = name,
                    Pos = fh.First,
                    Span = fh.Span,
                    Hits = fh.Positions,
                    HitKind = fh.Kind,
                    Run = fh.Run,
                    ScopeKind = rank,
                    IsBuiltin = isBuiltin,
                    RecentRank = Recent != null ? Recent.Rank(name) : 0,
                };
                c.Kind = Classify(c.Name);
                hits.Add(c);
            }

            if (hits.Count == 0) { Hide(); return false; }

            hits.Sort(Cmp);
            // 完全相同的名字只留一条（工程名与内建表可能重名），保留排在前面的那条
            var dedup = new List<Candidate>(hits.Count);
            var seenName = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in hits) if (seenName.Add(h.Name)) dedup.Add(h);
            hits = dedup;

            _wordComplete = word.Length > 0 &&
                            Rules.PrefixIsResolved(word, hits.Select(x => x.Name).ToList());

            if (!manual && HideOnComplete && _wordComplete)
            {
                // 打全了：候选一并不留 —— 否则按 Tab 会把已经打完的名字再补一遍
                Hide();
                return false;
            }

            _items = hits;
            _sel = 0;
            _top = 0;
            Visible = true;
            OnShow?.Invoke(_items, _sel, _top);
            return true;
        }

        /// <summary>光标所在的过程名（ScopeKey）；不在任何过程里返回 null。</summary>
        private string ProcOf(Snapshot snap)
        {
            if (snap == null) return null;
            foreach (var ms in _scopes)
            {
                if (!string.Equals(ms.Module, snap.Module, StringComparison.OrdinalIgnoreCase)) continue;
                var p = Parser.ProcOfLine(ms.Procs, snap.Line);
                return p == null ? null : p.ScopeKey;
            }
            return null;
        }

        /// <summary>
        /// `root.` 的成员名单；**类型不确定时返回 null**（调用方据此整个让给 VBE）。
        ///
        /// 类型来源只有两条，都是代码里写死的：`Dim c As Cls` / `Set c = New Cls`。
        /// 刻意不做推断（返回值类型、With 块、宿主对象）—— 猜错就会弹出
        /// "这个对象根本没有的成员"，比不弹更糟。
        /// </summary>
        private List<string> MemberPool(string root, Snapshot snap)
        {
            if (string.IsNullOrEmpty(root) || snap == null) return null;
            string curModule = snap.Module ?? "";
            string typeName;

            if (string.Equals(root, "Me", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(root, "this", StringComparison.OrdinalIgnoreCase))
            {
                typeName = curModule;
            }
            else
            {
                string proc = ProcOf(snap);
                typeName = null;
                foreach (var ms in _scopes)
                {
                    if (!string.Equals(ms.Module, curModule, StringComparison.OrdinalIgnoreCase)) continue;
                    typeName = ms.TypeOf(root, proc);
                    break;
                }
            }
            if (string.IsNullOrEmpty(typeName)) return null;

            var outp = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1) 同名模块（类模块 / 窗体 / 标准模块）的 Public 成员
            foreach (var ms in _scopes)
            {
                if (!string.Equals(ms.Module, typeName, StringComparison.OrdinalIgnoreCase)) continue;
                foreach (var kv in ms.Exported()) if (seen.Add(kv.Key)) outp.Add(kv.Key);
                return outp.Count > 0 ? outp : null;
            }

            // 2) `Type Point2` 里的成员（挂的是 "type:Point2" 伪过程键）
            foreach (var ms in _scopes)
                foreach (var d in ms.Decls)
                {
                    if (d.Proc == null || !d.Proc.StartsWith("type:", StringComparison.Ordinal)) continue;
                    if (!string.Equals(d.Proc.Substring(5), typeName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (seen.Add(d.Name)) outp.Add(d.Name);
                }

            return outp.Count > 0 ? outp : null;
        }

        /// <summary>"这个词是不是我们刚按 Tab 写进去的"。对不上就当场作废。</summary>
        private bool QuietFromAccepted(string word, Snapshot snap)
        {
            if (_acceptedWord == null) return false;
            if (Environment.TickCount64 - _acceptedTicks > AcceptQuietMs)
            {
                _acceptedWord = null;
                return false;
            }
            bool same = string.Equals(word, _acceptedWord, StringComparison.OrdinalIgnoreCase)
                        && snap.Line == _acceptedLine
                        && string.Equals(snap.Module ?? "", _acceptedModule ?? "",
                                         StringComparison.OrdinalIgnoreCase);
            if (same) return true;
            // 用户动过手（换行 / 换模块 / 改词）=> 记忆作废，下次照常提示
            _acceptedWord = null;
            return false;
        }

        /// <summary>
        /// ★排序键（全部"越大越靠前"，与 Python 版逐项对齐）：
        ///   命中质量(kind) > 首个命中位置 > 命中跨度
        ///     > 作用域就近 > 最近使用 > 你的名字(非内建) > 名字长度(短在前)
        /// </summary>
        public static int Cmp(Candidate a, Candidate b)
        {
            int c = b.HitKind.CompareTo(a.HitKind); if (c != 0) return c;
            c = a.Pos.CompareTo(b.Pos); if (c != 0) return c;
            c = a.Span.CompareTo(b.Span); if (c != 0) return c;
            c = b.ScopeKind.CompareTo(a.ScopeKind); if (c != 0) return c;
            c = b.RecentRank.CompareTo(a.RecentRank); if (c != 0) return c;
            c = a.IsBuiltin.CompareTo(b.IsBuiltin); if (c != 0) return c;
            c = a.Name.Length.CompareTo(b.Name.Length); if (c != 0) return c;
            return string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 图标档位。判据只有这一处：工程过程 > 工程类型 > 内建函数 > 内建类型 >
        /// 关键字 > 其余。拿不到档位就返回 null（不画图标），绝不默认画成"变量"。
        /// </summary>
        public string Classify(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            string kind;
            if (_procKinds.TryGetValue(name, out kind))
                return kind == "sub" ? "func" : (kind == "property" ? "prop" : "func");
            if (_typeNames.Contains(name)) return "type";

            if (Bank != null)
            {
                string bk = Bank.KindOf(name);
                if (bk == "sub" || bk == "function") return "func";
                if (bk == "property") return "prop";
                if (bk == "type") return "type";
            }

            if (VbaBuiltins.IsBuiltinFunction(name)) return "func";
            if (VbaBuiltins.IsBuiltinType(name)) return "type";
            if (VbaBuiltins.IsKeyword(name)) return "kw";
            return "var";
        }

        /// <summary>
        /// 按 Tab 确认之后，后面该跟什么：(补括号, 补空格)。
        /// 用户口径：Sub 补空格 / Function 补括号 / Property 都不补；
        /// 例外：在 Function 体内给自己赋值返回值时不补括号。
        /// </summary>
        public void Followup(string name, string enclosingProc, out bool wantParens, out bool wantSpace)
        {
            wantParens = false;
            wantSpace = false;
            if (string.IsNullOrEmpty(name)) return;
            string low = name.ToLowerInvariant();

            string kind;
            if (_procKinds.TryGetValue(name, out kind))
            {
                if (kind == "sub") { wantSpace = true; return; }
                if (kind == "function")
                {
                    if (IsOwnProcName(low, enclosingProc)) return;   // 给返回值赋值，不补
                    wantParens = true;
                    return;
                }
                return;                                              // property：都不补
            }

            if (VbaBuiltins.IsCallableBuiltin(name)) { wantParens = true; return; }

            bool? ov = Bank?.WantsSpace(name);
            wantSpace = ov ?? VbaBuiltins.KeywordWantsSpace(name);
        }

        public static bool IsOwnProcName(string name, string enclosingProc)
        {
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(enclosingProc)) return false;
            string b = enclosingProc.ToLowerInvariant();
            int hash = b.IndexOf('#');
            if (hash >= 0) b = b.Substring(0, hash);
            return name.ToLowerInvariant() == b;
        }

        // ------------------------------------------------------------------
        // 选择
        // ------------------------------------------------------------------

        public void Move(int delta)
        {
            if (!Visible || _items.Count == 0) return;
            _sel = (_sel + delta) % _items.Count;
            if (_sel < 0) _sel += _items.Count;
            ScrollToSelected();
            OnMove?.Invoke();
        }

        public void SetSelected(int index)
        {
            if (_items.Count == 0) return;
            _sel = Math.Max(0, Math.Min(_items.Count - 1, index));
            ScrollToSelected();
            OnMove?.Invoke();
        }

        public int PageSize => Math.Max(1, ViewRows - 1);

        public void ScrollToSelected()
        {
            if (_sel < _top) _top = _sel;
            else if (_sel >= _top + ViewRows) _top = _sel - ViewRows + 1;
            int maxTop = Math.Max(0, _items.Count - ViewRows);
            if (_top > maxTop) _top = maxTop;
            if (_top < 0) _top = 0;
        }

        public void ScrollBy(int lines)
        {
            if (!Visible) return;
            int maxTop = Math.Max(0, _items.Count - ViewRows);
            _top = Math.Max(0, Math.Min(maxTop, _top + lines));
            OnMove?.Invoke();
        }

        /// <summary>滚轮 / 拖滚动条：直接指定顶行。</summary>
        public void SetTop(int top)
        {
            if (!Visible) return;
            int maxTop = Math.Max(0, _items.Count - ViewRows);
            _top = Math.Max(0, Math.Min(maxTop, top));
            if (_sel < _top) _sel = _top;
            else if (_sel >= _top + ViewRows) _sel = _top + ViewRows - 1;
            OnMove?.Invoke();
        }

        public void Hide()
        {
            if (!Visible) return;
            Visible = false;
            _items = new List<Candidate>();
            _sel = 0; _top = 0;
            _wordComplete = false;
            OnHide?.Invoke();
        }

        /// <summary>接受选中的候选：写回 VBE，然后收起窗口。</summary>
        public bool Accept()
        {
            if (!Visible || _items.Count == 0) return false;
            var cand = _items[_sel];
            return AcceptName(cand.Name);
        }

        public bool AcceptName(string name)
        {
            if (string.IsNullOrEmpty(name) || Backend == null) { Hide(); return false; }

            // ★写完之前重新取一次快照：候选项算出来之后用户可能又敲了字，
            //   拿旧的 word_start_col 去替换会把行改坏。
            var snap = Backend.Snapshot();
            if (snap == null) { Hide(); return false; }

            int wordStart;
            string acceptLine = snap.Text ?? "";
            int acceptCaret = Rules.CharColFromDisp(acceptLine, snap.CaretCol);
            Rules.WordBefore(acceptLine, acceptCaret, out wordStart);
            int endCol = acceptCaret;
            if (wordStart > endCol) wordStart = endCol;

            string enclosing = null;
            foreach (var ms in _scopes)
            {
                if (!string.Equals(ms.Module, snap.Module, StringComparison.OrdinalIgnoreCase)) continue;
                var p = Parser.ProcOfLine(ms.Procs, snap.Line);
                enclosing = p == null ? null : p.Name;
                break;
            }
            if (enclosing == null) enclosing = Backend.EnclosingProcName();

            Followup(name, enclosing, out bool wantParens, out bool wantSpace);
            bool ok = Backend.ApplyCompletion(snap.Line, wordStart, endCol, name,
                wantParens, wantSpace, enclosing);
            Hide();
            if (ok)
            {
                // 记一笔"刚写进去的是这个词" —— 轮询下一拍发现文档变了会再触发一次，
                // 那时光标处正是这个词，不挡住的话用户会看到自己刚确认的名字又被提示。
                _acceptedWord = name;
                _acceptedModule = snap.Module ?? "";
                _acceptedLine = snap.Line;
                _acceptedTicks = Environment.TickCount64;
                Recent?.Note(name); Recent?.Flush();
                InvalidatePool();
            }
            return ok;
        }

        /// <summary>Tab 被按住时"这个按键该不该由我们吃掉"。</summary>
        public bool ShouldTakeAccept() => Visible && _items.Count > 0;

        /// <summary>
        /// 空格确认：只在"光标前的词还没打全"时才接管。
        /// ★旧版只要选中项是关键字就吃 —— 于是 `Set w` 时选中 `With`，按空格
        ///   就把 `w` 变成 `With`，`Set ws = ...` 直接写坏。Python 版的口径是
        ///   "already complete 就不吃"。
        /// </summary>
        public bool CanSpaceConfirm()
        {
            if (!Visible || _items.Count == 0) return false;
            if (_wordComplete) return false;
            var c = _items[_sel];
            return VbaBuiltins.IsKeyword(c.Name);
        }
    }
}
