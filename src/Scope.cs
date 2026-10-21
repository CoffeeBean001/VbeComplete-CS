// 作用域分析：从一段模块文本里算出"每个名字属于谁、在别处看不看得见"。
//
// ★为什么必须有这一层：旧的候选池是"整模块的全量标识符"，于是
//     Sub A 里的 `Dim tmp As Long`
//   会在
//     Sub B 里被提示出来 —— 而 VBA 里 B 根本看不见 A 的局部变量。
//   Python 版早就按 4/3/2 三档过滤了（本过程局部 > 本模块模块级 > 别的模块
//   Public），C# 版一直是"一锅端"，这是"提示里混着用不了的名字"的主因。
//
// ★这一层不碰 COM、不碰 UI，只吃字符串 —— 因此能把每一条可见性规则写成断言。
//
// 档位（越大越"近"，排序时优先）：
//   4 = 本过程的局部变量 / 参数
//   3 = 本模块的模块级名字（含 Private —— 自己模块当然看得见）
//   2 = 别的模块的 Public（以及"本模块引用过、但没声明"的名字）
//   0 = 不可见（别的过程的局部、别的模块的 Private）—— 不进候选池

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace VbeComplete
{
    internal sealed class Decl
    {
        public string Name;
        public string Proc;        // 过程的 ScopeKey；null = 模块级
        public bool IsPrivate;     // 只对模块级有意义
        public string KindHint;    // "sub" / "function" / "property" / "type" / "enum" / null
        public string TypeName;    // `As Cls` 里的 Cls（拿不到就是 null）
        public override string ToString() => $"{Name}@{Proc ?? "<module>"}{(IsPrivate ? "(p)" : "")}";
    }

    internal sealed class ModuleScope
    {
        public string Module;
        public List<ProcInfo> Procs = new List<ProcInfo>();
        public Dictionary<string, string> ProcKinds =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> TypeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public List<Decl> Decls = new List<Decl>();
        /// <summary>引用过、但没在本模块声明过的名字（隐式变量 / 宿主对象 / 控件）。
        /// 档位给 2：能用，但不如真声明过的名字靠前。
        /// ★模块里有 `Option Explicit` 时这一组会大幅收窄（见 <see cref="HostObjects"/>）。</summary>
        public List<string> FreeIdents = new List<string>();

        /// <summary>
        /// 只有【像宿主对象的自由标识符】才留在这里：出现过 `名字.` 或 `Set 名字 =`。
        /// 它们在 Option Explicit 的模块里也是合法的（窗体控件、Application、
        /// 隐式声明的对象），凭"代码里没写 Dim"就剔除会把 UserForm 上的按钮名
        /// 全删掉。
        /// </summary>
        public HashSet<string> HostObjects = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>模块头有没有 `Option Explicit`（= 强制声明变量）。</summary>
        public bool OptionExplicit;

        /// <summary>
        /// VBE 组件类型：1=标准模块、2=类模块、3=窗体、100=文档模块（ThisWorkbook /
        /// Sheet1）、0=未知（合成测试文本 / 还没读出来）。
        ///
        /// ★它只影响一件事：**这个模块的声明能不能进"裸名池"**。
        ///   类模块（2）/ 窗体（3）/ 文档模块（100）里的 `Public s1` 是**对象成员**，
        ///   只能 `c.s1` 这样访问；把它们并进裸名池，用户在别的模块里打一个 `s`
        ///   就会看到 `s1` —— 那是真·变量泄露（用户 v137 报的）。
        ///   当前正在编辑的那个模块不受这条限制（在自己模块里它当然是裸名）。
        /// </summary>
        public int ComponentType;

        /// <summary>行数。池子缓存用它判"文本变了没有"。</summary>
        public int LineCount;

        /// <summary>
        /// 变量名 -> 类型名（`Dim c As Cls` / `Set c = New Cls`）。
        /// key 是 "过程键|变量名"，模块级的键前缀为空。**只用来给 `c.` 找成员** ——
        /// 拿不到类型就什么都不做，绝不猜。
        /// </summary>
        public readonly Dictionary<string, string> VarTypes =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public string TypeOf(string name, string proc)
        {
            if (string.IsNullOrEmpty(name)) return null;
            string v;
            if (!string.IsNullOrEmpty(proc) && VarTypes.TryGetValue(proc + "|" + name, out v)) return v;
            if (VarTypes.TryGetValue("|" + name, out v)) return v;
            return null;
        }

        /// <summary>把本模块的声明按"当前过程"折算成 {小写名: (显示名, 档位)}。</summary>
        public Dictionary<string, KeyValuePair<string, int>> Visible(string currentProc)
        {
            var map = new Dictionary<string, KeyValuePair<string, int>>(StringComparer.OrdinalIgnoreCase);
            string cur = string.IsNullOrEmpty(currentProc) ? null : currentProc;

            foreach (var d in Decls)
            {
                if (d.Proc != null && !string.Equals(d.Proc, cur, StringComparison.OrdinalIgnoreCase))
                    continue;                       // 别的过程的局部：看不见
                int rank = d.Proc != null ? 4 : 3;
                Put(map, d.Name, rank);
            }
            foreach (var n in FreeIdents) Put(map, n, 2);
            return map;
        }

        /// <summary>只取"别的模块看得见"的那部分（本模块级、且不是 Private）。</summary>
        public Dictionary<string, KeyValuePair<string, int>> Exported()
        {
            var map = new Dictionary<string, KeyValuePair<string, int>>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in Decls)
            {
                if (d.Proc != null) continue;       // 局部一律不导出
                if (d.IsPrivate) continue;          // Private 在别的模块看不见
                Put(map, d.Name, 2);
            }
            return map;
        }

        private static void Put(Dictionary<string, KeyValuePair<string, int>> map, string name, int rank)
        {
            if (string.IsNullOrEmpty(name)) return;
            KeyValuePair<string, int> old;
            if (map.TryGetValue(name, out old) && old.Value >= rank) return;
            map[name] = new KeyValuePair<string, int>(name, rank);
        }
    }

    internal static class Scope
    {
        /// <summary>
        /// 有 `Option Explicit` 时要不要剔除未声明的隐式名字。
        /// 默认 1（按用户口径：强制声明的地方，没定义过的名字不该出现在候选里）；
        /// 设 0 可以退回旧行为（照收），留给"窗体控件名也想提示"的极端场景。
        /// </summary>
        private static readonly bool HonorExplicit = Settings.Flag("IMPLICIT_HONOR_EXPLICIT", true);

        // ★v133：标识符必须认 Unicode 字母 —— VBA 里 `学生 = Split(...)` / `总分()`
        //   是完全合法的，而旧正则只认 [A-Za-z_]，于是中文变量名从来没进过候选池，
        //   用户看到的就是"中文名一律不提示"。\p{L} 覆盖汉字/假名/谚文等，
        //   \p{N} 覆盖全角与阿拉伯数字。
        private const string Id = @"[\p{L}_][\p{L}\p{N}_]*[\%\&\#\@\$]?";

        private static readonly Regex ReProcHead = new Regex(
            @"^[ \t]*(?:(Public|Private|Friend|Global)[ \t]+)?(?:Static[ \t]+)?" +
            @"(Sub|Function|Property[ \t]+(?:Get|Let|Set)|Declare[ \t]+(?:PtrSafe[ \t]+)?(?:Sub|Function))" +
            @"[ \t]+(" + Id + ")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReEndProc = new Regex(
            @"^[ \t]*End[ \t]+(?:Sub|Function|Property)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReLocalDecl = new Regex(
            @"^[ \t]*(?:(?:Static)[ \t]+)?(?:Dim|Const|ReDim|WithEvents)[ \t]+(.*)$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // 过程外不带 Dim 的声明：`Private x As Long` / `Public g As Long`
        private static readonly Regex ReModuleBare = new Regex(
            @"^[ \t]*(Public|Private|Global|Friend)[ \t]+(?!(?:Sub|Function|Property|Declare|Type|Enum|Const|Dim|WithEvents|Event)\b)" +
            @"(" + Id + @")(?=[ \t]*(?:As\b|\(|$))",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // 过程外的 Dim / Const / WithEvents（修饰符可有可无）
        private static readonly Regex ReModuleDim = new Regex(
            @"^[ \t]*(?:(Public|Private|Global|Friend)[ \t]+)?(?:Dim|Const|WithEvents)[ \t]+(.*)$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReTypeHead = new Regex(
            @"^[ \t]*(?:(Public|Private)[ \t]+)?(Type|Enum)[ \t]+(" + Id + ")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReEndType = new Regex(
            @"^[ \t]*End[ \t]+(?:Type|Enum)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReIdent = new Regex(Id, RegexOptions.Compiled);

        // `Set c = New Cls` —— 没写 `Dim c As Cls` 时，这条是唯一的类型线索
        private static readonly Regex ReSetNew = new Regex(
            @"^[ \t]*Set[ \t]+(" + Id + @")[ \t]*=[ \t]*New[ \t]+(" + Id + ")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>解析一个模块。text 为 null 时返回一个空壳（绝不返回 null）。</summary>
        public static ModuleScope Analyze(string module, string text)
        {
            var ms = new ModuleScope { Module = module ?? "" };
            var lines = Parser.SplitLines(text);
            ms.LineCount = lines.Length;

            ms.Procs = Parser.ExtractProcedures(text);
            ms.ProcKinds = Parser.ProcKinds(ms.Procs);
            ms.TypeNames = Parser.ExtractTypeNames(text);

            // 行号 -> 所属过程（1-based）。过程头那一行算它自己。
            int n = lines.Length;
            var owner = new string[n + 2];
            foreach (var p in ms.Procs)
                for (int i = Math.Max(1, p.StartLine); i <= Math.Min(n, p.EndLine); i++)
                    if (owner[i] == null) owner[i] = p.ScopeKey;

            // 哪些名字是"某个过程的局部"—— 全量标识符收集时要排除它们，
            // 否则 B 里引用了 A 的局部名，就会把 A 的局部又带回来。
            var localNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            bool inTypeBlock = false;
            string typeOwner = null;
            bool typePrivate = false;
            bool typeIsEnum = false;

            for (int ln = 1; ln <= n; ln++)
            {
                string code = Parser.CodeOnly(lines[ln - 1]);
                if (code.Trim().Length == 0)
                {
                    // 空行/整行注释：不结束 Type 块（VBE 允许里面有空行）
                    continue;
                }

                // ---- Type / Enum 块 ----
                var tm = ReTypeHead.Match(code);
                if (tm.Success)
                {
                    inTypeBlock = true;
                    typeOwner = tm.Groups[3].Value;
                    typeIsEnum = tm.Groups[2].Value.Equals("Enum", StringComparison.OrdinalIgnoreCase);
                    typePrivate = tm.Groups[1].Success &&
                                  tm.Groups[1].Value.Equals("Private", StringComparison.OrdinalIgnoreCase);
                    ms.Decls.Add(new Decl
                    {
                        Name = typeOwner,
                        Proc = null,
                        IsPrivate = typePrivate,
                        KindHint = typeIsEnum ? "enum" : "type",
                    });
                    continue;
                }
                if (inTypeBlock)
                {
                    if (ReEndType.IsMatch(code)) { inTypeBlock = false; typeOwner = null; continue; }
                    // 成员：`Name As Long` 或 `Name = 1`
                    var mm = Regex.Match(code, @"^[ \t]*(" + Id + ")");
                    if (mm.Success)
                    {
                        string mem = Clean(mm.Groups[1].Value);
                        if (mem.Length > 0 && !VbaBuiltins.IsKeyword(mem))
                        {
                            // ★Enum 成员是**模块级可见**的（`clRed` 不用写 `Color2.clRed`）；
                            //   Type 成员只能通过变量访问（`p.X`），所以挂一个
                            //   "type:Owner" 的伪过程键，把它从普通候选里排除掉。
                            ms.Decls.Add(new Decl
                            {
                                Name = mem,
                                Proc = typeIsEnum ? null : "type:" + typeOwner,
                                IsPrivate = typePrivate,
                            });
                        }
                    }
                    continue;
                }

                string procKey = owner[ln];

                // ---- 过程头：参数是局部 ----
                var pm = ReProcHead.Match(code);
                if (pm.Success)
                {
                    string pname = Clean(pm.Groups[3].Value);
                    bool isPriv = pm.Groups[1].Success &&
                                  pm.Groups[1].Value.Equals("Private", StringComparison.OrdinalIgnoreCase);
                    string kind = "sub";
                    string head = pm.Groups[2].Value.ToUpperInvariant();
                    if (head.StartsWith("FUNCTION") || head.Contains("FUNCTION")) kind = "function";
                    else if (head.StartsWith("PROPERTY")) kind = "property";
                    else if (head.StartsWith("DECLARE")) kind = head.Contains("FUNCTION") ? "function" : "sub";

                    if (pname.Length > 0)
                        ms.Decls.Add(new Decl
                        {
                            Name = pname, Proc = null, IsPrivate = isPriv, KindHint = kind,
                        });

                    int lp = code.IndexOf('(');
                    int rp = code.LastIndexOf(')');
                    if (lp >= 0 && rp > lp)
                        foreach (var a in ParamTypes(code.Substring(lp + 1, rp - lp - 1)))
                        {
                            ms.Decls.Add(new Decl
                            {
                                Name = a.Key, Proc = procKey ?? pname, TypeName = a.Value,
                            });
                            if (a.Value != null)
                                ms.VarTypes[(procKey ?? pname) + "|" + a.Key] = a.Value;
                            localNames.Add(a.Key);
                        }
                    continue;
                }

                if (ReEndProc.IsMatch(code)) continue;
                if (Regex.IsMatch(code, @"^[ \t]*(?:Option|Attribute|#)", RegexOptions.IgnoreCase))
                    continue;

                // `Set c = New Cls`：给成员提示留一条类型线索
                var sn = ReSetNew.Match(code);
                if (sn.Success)
                {
                    string vn = Clean(sn.Groups[1].Value);
                    string tn = Clean(sn.Groups[2].Value);
                    if (vn.Length > 0 && tn.Length > 0)
                        ms.VarTypes[(procKey ?? "") + "|" + vn] = tn;
                    continue;
                }

                // ---- 过程内声明 ----
                if (procKey != null)
                {
                    var lm = ReLocalDecl.Match(code);
                    if (lm.Success)
                        foreach (var v in DeclTypes(lm.Groups[1].Value))
                        {
                            ms.Decls.Add(new Decl { Name = v.Key, Proc = procKey, TypeName = v.Value });
                            if (v.Value != null) ms.VarTypes[procKey + "|" + v.Key] = v.Value;
                            localNames.Add(v.Key);
                        }
                    continue;
                }

                // ---- 过程外（模块级）----
                var dm = ReModuleDim.Match(code);
                if (dm.Success)
                {
                    string mod = dm.Groups[1].Success ? dm.Groups[1].Value : "";
                    bool isPriv = mod.Equals("Private", StringComparison.OrdinalIgnoreCase);
                    foreach (var v in DeclTypes(dm.Groups[2].Value))
                    {
                        ms.Decls.Add(new Decl { Name = v.Key, Proc = null, IsPrivate = isPriv, TypeName = v.Value });
                        if (v.Value != null) ms.VarTypes["|" + v.Key] = v.Value;
                    }
                    continue;
                }

                var bm = ReModuleBare.Match(code);
                if (bm.Success)
                {
                    string v = Clean(bm.Groups[2].Value);
                    bool isPriv = bm.Groups[1].Value.Equals("Private", StringComparison.OrdinalIgnoreCase);
                    if (v.Length > 0 && !VbaBuiltins.IsKeyword(v))
                    {
                        var tm2 = Regex.Match(code, @"\bAs[ \t]+(?:New[ \t]+)?(" + Id + ")",
                            RegexOptions.IgnoreCase);
                        string tn = tm2.Success ? Clean(tm2.Groups[1].Value) : null;
                        ms.Decls.Add(new Decl { Name = v, Proc = null, IsPrivate = isPriv, TypeName = tn });
                        if (tn != null) ms.VarTypes["|" + v] = tn;
                    }
                    continue;
                }
            }

            ms.OptionExplicit = Parser.OptionExplicit(text);

            // ---- 全量标识符：只留"没被任何过程声明为局部"的 ----
            var declared = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in ms.Decls) declared.Add(d.Name);

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var raw in lines)
            {
                string code = Parser.CodeOnly(raw);
                if (code.Length == 0) continue;
                foreach (Match m in ReIdent.Matches(code))
                {
                    string w = Clean(m.Value);
                    if (w.Length == 0) continue;

                    // ★v146：**前面紧挨着点号 = 成员访问**（`wsR.Cells` 里的 Cells、
                    //   `Range.Value` 里的 Value、`Debug.Print` 里的 Print）—— 它不是
                    //   "本模块的自由名字"，收进来就会被当成全工程可用的裸名：
                    //   用户在任意模块打 `cell` 都会冒出 `Cells`。这就是 v146 用户
                    //   报的"输入 cell 提示 cells，是不是泄露"—— 确实是泄露，
                    //   泄露的是**别的模块里某个对象的成员名**。
                    //   注意方向别搞反：`LooksLikeHostObject` 看的是"**后面**跟点号"
                    //   （`CommandButton1.Caption` → 对象名要留），这里看的是"**前面**
                    //   是点号"（→ 成员名要丢）。
                    //   裸用 `Cells(1,1)` 仍然照收 —— 那是用户真这么写了，该提示。
                    if (IsMemberAccess(code, m)) continue;

                    if (VbaBuiltins.IsKeyword(w)) continue;
                    if (declared.Contains(w)) continue;
                    if (localNames.Contains(w)) continue;             // 别处的局部，不能漏过来

                    // 单字母（且是 ASCII）噪音太大 —— `i` `n` 这种在候选里没意义。
                    // ★汉字单字不在此列：`和` / `总` 是正经名字，不能按字母数筛掉。
                    if (w.Length < 2 && w[0] < 0x80) continue;

                    bool host = LooksLikeHostObject(code, m);
                    if (host) ms.HostObjects.Add(w);

                    // ★强制声明变量的模块：没写 Dim 的名字在 VBE 里直接编译不过，
                    //   把它当候选等于教用户写错代码。唯一例外是"像宿主对象"的
                    //   那些（控件名 / Application / 隐式对象）—— 它们本来就不用声明。
                    if (HonorExplicit && ms.OptionExplicit && !host) continue;

                    if (seen.Add(w)) ms.FreeIdents.Add(w);
                }
            }
            return ms;
        }

        /// <summary>
        /// 这个标识符是不是"点号后面的成员名"（`wsR.Cells` 里的 `Cells`）？
        /// 只看紧邻的前一个非空白字符：是 `.` 就算成员，否则不算。
        /// 不追类型、不认大小写 —— 判据纯语法，绝不猜。
        /// </summary>
        private static bool IsMemberAccess(string code, Match m)
        {
            int p = m.Index - 1;
            while (p >= 0 && (code[p] == ' ' || code[p] == '\t')) p--;
            return p >= 0 && code[p] == '.';
        }

        /// <summary>
        /// 这个名字像不像"宿主给的对象"（窗体控件、Application、隐式全局）？
        /// 两条判据，都只认最直接的语法证据：
        ///   * 后面紧跟点号 —— `CommandButton1.Caption`
        ///   * 前面是 `Set` —— `Set btn = ...`
        /// 不做类型推断：拿不准就按普通自由标识符处理。
        /// </summary>
        private static bool LooksLikeHostObject(string code, Match m)
        {
            int after = m.Index + m.Length;
            while (after < code.Length && (code[after] == ' ' || code[after] == '\t')) after++;
            if (after < code.Length && code[after] == '.') return true;

            int p = m.Index - 1;
            while (p >= 0 && (code[p] == ' ' || code[p] == '\t')) p--;
            if (p >= 2)
            {
                string head = code.Substring(0, p + 1).TrimEnd();
                if (head.Length >= 3 &&
                    head.Substring(head.Length - 3).Equals("Set", StringComparison.OrdinalIgnoreCase) &&
                    (head.Length == 3 || !char.IsLetterOrDigit(head[head.Length - 4])))
                    return true;
            }
            return false;
        }

        /// <summary>`a, b(1 To 3) As Long` -> [(a, Long), (b, Long)]</summary>
        private static IEnumerable<KeyValuePair<string, string>> DeclTypes(string rest)
        {
            foreach (var part in Parser.SplitTopLevel(rest))
            {
                string p = part.Trim();
                if (p.Length == 0) continue;
                // `x As New Cls` 里的 As 之后不参与；只取首个标识符
                var mm = Regex.Match(p, "^(" + Id + ")");
                if (!mm.Success) continue;
                string name = Clean(mm.Groups[1].Value);
                if (name.Length == 0 || VbaBuiltins.IsKeyword(name)) continue;

                string type = null;
                var tm = Regex.Match(p, @"\bAs[ \t]+(?:New[ \t]+)?(" + Id + ")",
                    RegexOptions.IgnoreCase);
                if (tm.Success)
                {
                    string t = Clean(tm.Groups[1].Value);
                    if (t.Length > 0) type = t;
                }
                yield return new KeyValuePair<string, string>(name, type);
            }
        }

        private static IEnumerable<KeyValuePair<string, string>> ParamTypes(string inner)
        {
            foreach (var part in Parser.SplitTopLevel(inner))
            {
                string p = Regex.Replace(part, @"^\s*(?:Optional|ByVal|ByRef|ParamArray)\b",
                        "", RegexOptions.IgnoreCase).Trim();
                p = Regex.Replace(p, @"^\s*(?:Optional|ByVal|ByRef|ParamArray)\b",
                        "", RegexOptions.IgnoreCase).Trim();
                p = p.TrimStart('(');
                var mm = Regex.Match(p, "^(" + Id + ")");
                if (!mm.Success) continue;
                string name = Clean(mm.Groups[1].Value);
                if (name.Length == 0 || VbaBuiltins.IsKeyword(name)) continue;

                string type = null;
                var tm = Regex.Match(p, @"\bAs[ \t]+(?:New[ \t]+)?(" + Id + ")",
                    RegexOptions.IgnoreCase);
                if (tm.Success)
                {
                    string t = Clean(tm.Groups[1].Value);
                    if (t.Length > 0) type = t;
                }
                yield return new KeyValuePair<string, string>(name, type);
            }
        }

        private static string Clean(string name)
        {
            if (string.IsNullOrEmpty(name)) return "";
            return name.TrimEnd('%', '&', '!', '#', '@', '$').Trim();
        }
    }
}
