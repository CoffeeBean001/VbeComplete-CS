// 候选引擎：攒池子、匹配、排序、决定弹不弹、选完之后补什么。
// ★不引用 WinForms（窗口靠回调接进来），所以"弹不弹 / 排第几 / 补什么"都能离线测。
// ★排序键顺序对齐 Python 版（用户在真机上调出来的）：命中质量 > 位置 > 跨度 > 作用域 >
//   最近使用 > 内建 > 名字长度。⚠️ 别把"作用域"提前，它会顶掉匹配更好的候选。

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
        public int ScopeKind;      // 4 本过程局部 / 3 本模块 / 2 别的模块 Public
        public int RecentRank;     // 最近使用次数
        public bool IsBuiltin;
    }

    internal sealed class Engine
    {
        public IVbeSource Backend;
        public Recent Recent;

        // ★一屏最多几行：与 Python 版 MAX_VISIBLE_ROWS 对齐（15）。候选不足时
        //   窗口自己变矮 —— 用户口径是"显示几行就几行高"，不是固定 15 行留白。
        // ★v175：改成**现读**（原来是 readonly 快照）—— 设置界面改完要立刻生效，
        //   否则用户勾掉"打全即收"还得重启程序，看起来像没保存。
        //   每次弹窗读几次 Settings 是字典查表 + 两个环境变量，代价可以忽略。
        public int ViewRows => Settings.Int("VIEW_ROWS", 15, 3, 30);
        public bool HideOnComplete => Settings.Flag("HIDE_ON_COMPLETE", true);
        // ★v138：不是 readonly —— 探针要能当场切成 1 做 A/B（证明"点号后不弹"
        //   确实是这条判据在起作用），运行期只读一次 settings.ini。
        public bool MemberPopup = Settings.Flag("MEMBER_POPUP", false);
        public bool DeclPopup => Settings.Flag("DECL_POPUP", false);

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

        // 池：全工程的声明。★按行数缓存在 VbeBackend 那一层（ProjectScopes），
        //   这里只留当前这一轮的解析结果 —— 曾经这里还有一份 _scopeCache，
        //   只写不读，纯属"以为在用"（v150 删）。
        private List<ModuleScope> _scopes = new List<ModuleScope>();
        private List<string> _poolNames = new List<string>();
        // 池里哪些名字【必须】通过作用域判定才能出现（= 声明出来的名字）。
        // 内建 / 组件名 / 词库 / "引用过但没声明"这些是结构性的，不受作用域约束。
        private readonly HashSet<string> _poolScoped = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> _procKinds =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private HashSet<string> _typeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string _poolKey;

        // 刚按 Tab 确认过的那个词。同一行、同一个词不再提示自己 ——
        // Python 版 v90 的那条口径（"按完 Tab 意味着本次录入完成，不要再匹配提示"）。
        private string _acceptedWord;
        private string _acceptedModule;
        private int _acceptedLine;
        private long _acceptedTicks;
        private const long AcceptQuietMs = 2000;

        // 这一次触发时"光标前的词是不是已经打全了"（空格确认要用）
        private bool _wordComplete;

        // ★v157：让位给 VBE 成员列表时的宽限状态（键 = 模块|行|显示列）。
        private string _yieldKey;
        private long _yieldUntil;
        private const int YieldGraceMs = 250;

        // ★v151：这一轮候选是"哪个模块、哪一行、从哪个字符列起、什么词"算出来的。
        //   AcceptName 拿它复核：候选算出来之后用户可能按了 Home / End，或者
        //   用鼠标点到了同一行的别处 —— 行列变了但我们手里还是旧候选，
        //   照旧写进去就是"把候选插到行首"或者"把另一个词截成两半"。
        //   为 null 表示"没有上下文"（直接调 AcceptName 的探针走这条路）。
        private string _ctxModule;
        private int _ctxLine;
        private int _ctxWordStart;

        public IReadOnlyList<Candidate> Items => _items;
        public int SelectedIndex => _sel;
        public int TopIndex => _top;

        /// <summary>
        /// 上一次触发里被**截断掉**的命中总数（0 = 没截断）。
        /// 池子可能几百条，列表只留前 240 条 —— 这个数给自检与日志用。
        /// </summary>
        public int TruncatedFrom { get; private set; }

        /// <summary>当前裸名池的条目数（诊断用：池子空了会表现为"什么都不弹"）。</summary>
        public int PoolCount => _poolNames.Count;

        // ------------------------------------------------------------------
        // 池子
        // ------------------------------------------------------------------

        /// <summary>
        /// 重建池子。key 带"本模块文本长度 + 全工程签名 + 被挖掉的词"：本模块边打字边
        /// 重解析（新写的变量立刻能提示），别的模块只在行数变了之后才重解析。
        /// blank*(1-based) 是正在输入的那个词 —— 解析前挖成空格，池子里就不会出现它自己。
        /// </summary>
        private void EnsurePool(string moduleName, string moduleText,
            int blankLine = 0, int blankStart = 0, string blankToken = null)
        {
            int sig = 0;
            List<ModuleScope> others = null;
            if (Backend != null)
            {
                others = Backend.ProjectScopes(out sig);
            }

            string key = (moduleName ?? "") + "|" + (moduleText == null ? -1 : moduleText.Length) +
                         "|" + sig + "|" + blankLine + ":" + blankStart + ":" + (blankToken ?? "");
            if (_poolNames.Count > 0 && _poolKey == key) return;

            var scopes = new List<ModuleScope>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _procKinds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _typeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1) 当前模块：用传进来的文本（最新鲜），不吃缓存；先把半截词挖掉
            ModuleScope cur = null;
            if (moduleText != null)
            {
                string parsed = !string.IsNullOrEmpty(blankToken)
                    ? Rules.BlankToken(moduleText, blankLine, blankStart, blankToken)
                    : moduleText;
                cur = Scope.Analyze(moduleName ?? "", parsed);
                // ★v174：Analyze 出来的新对象没有组件类型（默认 0 = 未知），而"这个模块是
                //   标准模块还是类模块"决定了 Property / Me 该不该提示。缓存里同名的那个
                //   是 ProjectScopes 给的，类型已经填好 —— 抄过来。抄不到就保持未知（不过滤）。
                if (others != null)
                    foreach (var ms in others)
                        if (string.Equals(ms.Module, moduleName ?? "",
                                          StringComparison.OrdinalIgnoreCase))
                        {
                            cur.ComponentType = ms.ComponentType;
                            // ★v183：窗体控件名也不在文本里（见 Scope.ModuleScope.Controls），
                            //   同样只能从缓存里那份抄过来 —— 不抄的话池子里只有一个"空窗体"。
                            if (ms.Controls != null && ms.Controls.Count > 0)
                                cur.Controls = ms.Controls;
                            break;
                        }
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
                // ★对象模块（类模块 2 / 窗体 3 / 文档模块 100）的声明是**成员**，只能
                //   `c.s1` 这样访问，绝不能当裸名提示（v137 报的"变量泄露"）。正在编辑的
                //   那个模块例外；类型未知（0）按标准模块处理，宁可多提示。
                bool isCur = ReferenceEquals(ms, cur) ||
                             string.Equals(ms.Module, moduleName ?? "",
                                           StringComparison.OrdinalIgnoreCase);
                if (!isCur && ms.ComponentType != 0 && ms.ComponentType != 1) continue;

                // ★v183：窗体控件名。控件是**本窗体的模块级名字** —— 在窗体自己的代码里
                //   裸名直接可用（`TextBox1.Text = "x"`），在别的模块里必须写
                //   `UserForm1.TextBox1`。所以只收当前模块的：收别的窗体的控件名，它永远
                //   不会在这个模块里可见，只是白占池子、把用户的候选挤下去。
                //   放在 Decls 之前：控件是用户最想补全的那一类，同分时靠前更顺手。
                if (isCur && ms.Controls != null && !Settings.Flag("NO_FORM_CTL", false))
                    foreach (var n in ms.Controls)
                    {
                        if (nameSeen.Add(n)) names.Add(n);
                        _poolScoped.Add(n);
                    }

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
                foreach (var f in ms.FreeIdents)
                {
                    if (nameSeen.Add(f)) names.Add(f);
                    // ★v154：无 Option Explicit 时"用过但没声明"的名字在 VBA 里是模块私有。
                    //   不进 _poolScoped 的话，A 模块的 `foo = 1` 会让 B 模块打 `fo` 也冒出它。
                    _poolScoped.Add(f);
                }
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

            // 5) ★词库：keywords-cs.db + functions-cs.db 里的全部词条。
            //   旧版**根本没有把内建名字放进池子** —— 只有"工程里恰好用到过的"
            //   才会经由全量标识符混进来，于是 `MsgBox` 在没写过的模块里搜不到。
            //   现在两个库在启动时已经全量载入内存，这里直接并进池子。
            //   它们不受作用域约束（不进 _poolScoped），排序时靠 IsBuiltin 往后排。
            foreach (var n in VbaBuiltins.Pool())
                if (nameSeen.Add(n)) names.Add(n);

            _poolNames = names;
            _poolKey = key;
        }

        /// <summary>把每个名字折算成"作用域档位"，key 为小写名。</summary>
        private Dictionary<string, int> BuildRankMap(string curModule, string curProc)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var ms in _scopes)
            {
                bool isCur = string.Equals(ms.Module, curModule, StringComparison.OrdinalIgnoreCase);
                // ★v160：对象模块的成员也不能进 rank 表。否则同名的名字从别的路（比如别的
                //   模块的隐式变量 s1）进池后，会在这里被"认领"成 rank 2 并通过作用域判定
                //   —— 就是 v137 那个"模块1 打 s 冒出 cls1 的 s1"。
                if (!isCur && ms.ComponentType != 0 && ms.ComponentType != 1) continue;
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

            // VBE 自己的成员列表正画着 —— 一律让位：它的列表有真类型信息，比我们准，
            // 等它关掉之后下一次触发我们再补位。
            // ★v157：让位带 250ms 宽限期（它的列表要等 WM_TIMER 才 Show，我们往往先跑完，
            //   只判"此刻可见"会闪一下）。同一位置让过一次后 250ms 内继续让，之后照常弹。
            if (!manual)
            {
                string ykey = (snap.Module ?? "") + "|" + snap.Line + "|" + snap.CaretCol;
                if (YieldToVbe != null && YieldToVbe())
                {
                    _yieldKey = ykey;
                    _yieldUntil = Environment.TickCount64 + YieldGraceMs;
                    Hide();
                    return false;
                }
                if (string.Equals(_yieldKey, ykey, StringComparison.Ordinal) &&
                    Environment.TickCount64 < _yieldUntil)
                {
                    Hide();
                    return false;
                }
            }

            string line = snap.Text ?? "";
            // ★GetSelection 给的是**显示列**，后面所有取词/判据都按字符列算
            //   —— 含中文/全角的行不换算就会取到空词，连弹窗都不出现。
            int caret = Rules.CharColFromDisp(line, snap.CaretCol);

            // 注释（`'` 或 `Rem`）里 / 字符串里一律不弹 —— 判据见 Rules.InDeadZone。
            // ★v152（对标 IDEA）：手动 Ctrl+Space 也拦。IDEA 在注释/字符串里同样不给
            //   代码补全 —— 那里弹出的候选按 Tab 只会把注释或字符串内容改写掉，
            //   没有任何语义价值。以前 `!manual` 的豁免让这条路原样保留着。
            if (Rules.InDeadZone(line, caret)) { Hide(); return false; }
            // ★声明位置一律不弹（含 `As` 之后、裸声明 `Public x` 之后）。
            //   见 Rules.InDeclarationPosition 的说明。
            if (!manual && !DeclPopup && Rules.InDeclarationPosition(line, caret)) { Hide(); return false; }

            // 先取词：池子要拿"这个词被挖掉之后"的文本去解析。
            int wordStart;
            string word = Rules.WordBefore(line, caret, out wordStart);

            // ★`As` / `New` 之后一律不弹（用户 v137 口径）。
            //   与上面的 InDeclarationPosition 互补：那条管 `Dim x As |` 这种
            //   声明开头的整行，这条管 `Set c = New |` / `dic.Add 1, New |` /
            //   `For i As |` / 过程参数 `a As |` —— 行首都不是声明关键字，
            //   旧实现里它们会照常弹出全工程候选（用户报的第 2 条）。
            if (!manual && !DeclPopup && Rules.AfterTypeKeyword(line, wordStart))
            {
                Hide();
                return false;
            }

            // ★v138：点号之后**一个候选都不出**，整个让给 VBE（同 PY v103c，用户原话：
            //   "输入 student.，点号后面就不要出提示框"）。池子是裸名池，出在点号后面就是噪音。
            //   位置判据只有一份：Rules.MemberDotCol。
            //   ⚠️ 放在 EnsurePool 之前：点号是高频键，这一步不需要池子也不需要 COM。
            //   ⚠️ MEMBER_POPUP=1 恢复旧行为；Ctrl+Space（manual）两种情形照弹。
            if (!manual && !MemberPopup && Rules.MemberDotCol(line, caret) >= 0)
            {
                Hide();
                return false;
            }

            string text = Backend?.ModuleText();
            EnsurePool(snap.Module, text, snap.Line, wordStart, word);

            // ---- `对象.` 位置（只剩 MEMBER_POPUP=1 和 Ctrl+Space 走得到）----
            // 能定出 `root` 的声明类型就按类型收窄到真成员；定不出来就照旧出
            // 全工程候选（Python 版同一口径：宁可给宽的，也不给"查到了类型但
            // 一个成员都没收进来"的空窗）。
            List<string> memberPool = null;
            bool atMemberPos = false;
            if (Rules.MemberRoot(line, caret, out string root, out int dotCol))
            {
                atMemberPos = true;
                memberPool = MemberPool(root, snap);
            }

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

            // ★声明关键字后面（`Private Fu|` / `Preserve re|`）只留该位置合法的东西：
            //   关键字位置只出关键字（不然 `Private P|` 会把 Public / Print 一起塞进来），
            //   纯语法位置（End / Exit / On Error …）连名字也不合法。三态见 Rules.DeclSlotOf。
            string[] declFollow = null;
            bool declNamesBlocked = false;
            if (word.Length > 0)
            {
                // ★v180：形参表里"一个新参数的开头"是独立的一格 —— 那里要的是
                //   ByVal / ByRef / Optional / ParamArray（参数名照留）。以前它被
                //   "声明位置不弹"整条封死（用户报：形参列表里这两条一条都不提示）。
                declFollow = Rules.ParamStartFollow(line, wordStart);
                if (declFollow == null)
                {
                    var slot = Rules.DeclSlotOf(line, wordStart, out declFollow);
                    if (slot == Rules.DeclSlot.None) declFollow = null;
                    declNamesBlocked = slot == Rules.DeclSlot.KeywordsOnly;
                }
            }

            // ★v174：关键字的"使用场合"限制（用户报：标准模块里冒出 Property、
            //   函数体里冒出 Private）。两个事实就够：这个模块是什么种类、
            //   光标在不在某个过程体内。拿不准时（类型未知 / 解析不到过程）一律放行。
            string curProc = ProcOf(snap);
            var ranks = BuildRankMap(snap.Module, curProc);
            var modKind = Rules.KindOfComponent(ComponentTypeOf(snap.Module));
            bool inProc = curProc != null;

            var hits = new List<Candidate>(64);
            foreach (var name in memberPool ?? _poolNames)
            {
                if (string.IsNullOrEmpty(name)) continue;
                FuzzyHit fh;
                if (!Fuzzy.Match(name, word, out fh)) continue;

                // 单字母只认前缀/精确 —— 否则打一个 `e` 会命中几百个
                if (word.Length == 1 && fh.Kind < 3) continue;

                bool isBuiltin = VbaBuiltins.IsBuiltin(name);

                // ★v174：关键字只在它合法的场合出现。`Property` 进不了标准模块，
                //   `Private` / `Public` 进不了过程体，`ReDim` 出不了过程体 —— 这些位置
                //   按 Tab 补进去就是一行红的代码，不如一条都不给。手动 Ctrl+Space 同样拦
                //   （同 v152 的口径：位置根本不合法，多给一条只是多一次踩坑机会）。
                if (isBuiltin && !Rules.KeywordAllowed(name, modKind, inProc)) continue;

                // 声明关键字后面的位置：把"这个位置根本不可能出现的东西"剔掉。
                // ★v170：判据从 IsKeyword 放宽到 IsBuiltin —— `Reset` / `Unload` / `Round`
                //   这类词不在关键字库里（在语句/函数表里），打 `ReDim Preserve res|`
                //   会把 `Reset` 塞进来，正是用户说的"这里不该出东西"。
                // ★v173：纯语法位置连**名字**一起剔（`Preserve re|` 只该有 ReDim）。
                if (declFollow != null && !manual)
                {
                    if (isBuiltin && !Rules.InList(declFollow, name)) continue;
                    if (!isBuiltin && declNamesBlocked) continue;
                }

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
                    // 声明出来的名字查不到档位 = 当前位置看不见它（别的过程的局部 / 别的
                    // 模块的 Private），直接剔除。
                    // ★v151：内建是全局可见的 —— 用户在某过程里 `Dim Left As Long`，别的过程
                    //   里内建的 Left 就会被这条一票否决（词库那份已因同名去重没进池）。
                    if (!isBuiltin && _poolScoped.Contains(name)) continue;
                    rank = 2;
                }

                var c = new Candidate
                {
                    Name = name,
                    Pos = fh.First,
                    Span = fh.Span,
                    Hits = fh.Positions,
                    HitKind = fh.Kind,
                    ScopeKind = rank,
                    IsBuiltin = isBuiltin,
                    RecentRank = Recent != null ? Recent.Rank(name) : 0,
                };
                c.Kind = Classify(c.Name);
                hits.Add(c);
            }

            if (hits.Count == 0) { Hide(); return false; }

            hits.Sort(Cmp);

            // ★v152（对标 IDEA）：命中太多就截断。池子里光词库就 261 条，打一两个
            //   字母很容易命中几百个 —— 滚动条长得没有意义，每次重画也白费。
            //   排序已经保证"最像的"在前 240 条里，继续输入一个字母就能收窄。
            //   ★截断只发生在**给 UI 的列表**上；自检要看全部命中就看 TruncatedFrom。
            const int MaxCandidates = 240;
            TruncatedFrom = hits.Count > MaxCandidates ? hits.Count : 0;
            if (hits.Count > MaxCandidates)
                hits.RemoveRange(MaxCandidates, hits.Count - MaxCandidates);
            // 完全相同的名字只留一条（工程名与内建表可能重名），保留排在前面的那条
            var dedup = new List<Candidate>(hits.Count);
            var seenName = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in hits) if (seenName.Add(h.Name)) dedup.Add(h);
            hits = dedup;

            // ★单字符输入**永远不算打全**：`For` 后面的位置把内建全剔了，候选只剩 `i` 一条
            //   ⇒ PrefixIsResolved 认为"打全了" ⇒ 窗口刚算出来就被 HideOnComplete 收掉，
            //   用户看到的是"没提示"。打一个字母还没到"确认完毕"的阶段。
            _wordComplete = word.Length > 1 &&
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
            // 记下"这批候选属于哪个位置"，Accept 时复核（见 AcceptName）。
            _ctxModule = snap.Module ?? "";
            _ctxLine = snap.Line;
            _ctxWordStart = wordStart;
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

        /// <summary>这个模块的 VBE 组件类型（0 = 不知道）。见 Rules.KindOfComponent。</summary>
        private int ComponentTypeOf(string module)
        {
            foreach (var ms in _scopes)
                if (string.Equals(ms.Module, module ?? "", StringComparison.OrdinalIgnoreCase))
                    return ms.ComponentType;
            return 0;
        }

        /// <summary>
        /// `root.` 的成员名单；**类型不确定时返回 null**（调用方据此整个让给 VBE）。
        /// 类型来源只有 `Dim c As Cls` / `Set c = New Cls` 两条，刻意不做推断 ——
        /// 猜错就会弹出"这个对象根本没有的成员"，比不弹更糟。
        /// </summary>
        private List<string> MemberPool(string root, Snapshot snap)
        {
            if (string.IsNullOrEmpty(root) || snap == null) return null;
            string curModule = snap.Module ?? "";
            string typeName;
            // ★v150：`Me.` / `this.` 是"本模块自己"，成员含 Private；
            //   别的 root 是"别人的类型"，只能看 Public。
            bool self = string.Equals(root, "Me", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(root, "this", StringComparison.OrdinalIgnoreCase);

            if (self)
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

            // 1) 同名模块（类模块 / 窗体 / 标准模块）的成员：
            //    Me -> 含 Private；别的模块 -> 只取 Public。
            foreach (var ms in _scopes)
            {
                if (!string.Equals(ms.Module, typeName, StringComparison.OrdinalIgnoreCase)) continue;
                foreach (var kv in (self ? ms.Members() : ms.Exported()))
                    if (seen.Add(kv.Key)) outp.Add(kv.Key);
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
        /// 图标档位。判据只有这一处：工程过程 > 工程类型 > 词库函数 > 词库类型 >
        /// 词库关键字 > 其余。拿不到档位就返回 null（不画图标），绝不默认画成"变量"。
        /// </summary>
        public string Classify(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            string kind;
            if (_procKinds.TryGetValue(name, out kind))
                return kind == "sub" ? "func" : (kind == "property" ? "prop" : "func");
            if (_typeNames.Contains(name)) return "type";

            string bk = VbaBuiltins.KindOf(name);
            if (bk != null) return bk;

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

            // 关键字要不要空格、以及用户自己在词库里标了 wants_space 的那些
            wantSpace = VbaBuiltins.KeywordWantsSpace(name);
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
            // ★v152（对标 IDEA）：补全替换的是**整个标识符**，不是"光标左边那一截"。
            //   光标停在 `MsgB|ox` 中间按 Tab 时，旧实现只换 [wordStart, caret)，
            //   结果是 `MsgBoxox`。IDEA 在这里也是整词替换 —— 右侧连续的标识符字符
            //   一并吃掉（遇到 `(` `.` `,` 等就停）。
            while (endCol <= acceptLine.Length && Rules.IsIdentChar(acceptLine[endCol - 1]))
                endCol++;

            // ★v151：复核"现在还是当初算候选的那个位置吗"。候选窗开着时按 Home/End 或点到
            //   同一行别处，行文本没变、轮询不重触发；此时按 Tab 会拿新的 wordStart 去替换
            //   （插到行首 / 把词截成两半）。只比行号/取词起点/模块，不比词本身（多打一个
            //   字母是正常的，那时 wordStart 不变、必须放行）。
            if (_ctxModule != null &&
                (snap.Line != _ctxLine ||
                 wordStart != _ctxWordStart ||
                 !string.Equals(snap.Module ?? "", _ctxModule, StringComparison.OrdinalIgnoreCase)))
            {
                Hide();
                return false;
            }

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
                // ★v151：这里只记一笔，**不**立刻写盘 —— 见 App.PollOnce 里的合并 Flush。
                //   原来每接受一次候选就 WriteAllText + File.Move 一次，家目录在
                //   OneDrive / 网络盘时确认完会明显卡一下，而且记录只增不减。
                Recent?.Note(name);
                InvalidatePool();
            }
            return ok;
        }

    }
}
