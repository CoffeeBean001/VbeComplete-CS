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
        /// ★模块里有 `Option Explicit` 时这一组会大幅收窄：出现过 `名字.` 或
        /// `Set 名字 =` 的（窗体控件 / Application / 隐式对象）仍然保留，
        /// 凭"代码里没写 Dim"就剔除会把 UserForm 上的按钮名全删掉。</summary>
        public List<string> FreeIdents = new List<string>();

        /// <summary>
        /// 隐式变量**首次出现**所在的过程（ScopeKey；null = 模块级 / 过程外）。
        /// <para>★v155（对齐 PY 版默认的 scope="proc"）：VBA 里未声明变量的作用域
        /// 就是它首次出现的过程 —— A 过程里 `num = 1`、B 过程里也写 `num`，那是
        /// 两个变量。把 A 的那个提示到 B 去，等于教用户写一行运行时报错的代码。</para>
        /// </summary>
        public readonly Dictionary<string, string> FreeIdentProc =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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
            foreach (var n in FreeIdents)
            {
                // 隐式变量只在"它首次出现的过程"里可见（过程外的那些全模块可见）
                string p;
                if (FreeIdentProc.TryGetValue(n, out p) && !string.IsNullOrEmpty(p) &&
                    !string.Equals(p, cur, StringComparison.OrdinalIgnoreCase))
                    continue;
                Put(map, n, 2);
            }
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

        /// <summary>
        /// 本模块自己的**全部模块级成员**（含 Private），给 `Me.` 用。
        ///
        /// ★v150：`Me.` 后面 Private 成员完全合法（窗体里的 Private Sub、
        ///   类模块里的 Private m_x），拿 <see cref="Exported"/> 去填是错的 ——
        ///   那份名单是"别的模块看得见什么"，正好把 Private 全滤掉了，
        ///   于是 `Me.` 后面收不到自己的私有过程。
        /// ★不收过程的局部（`Proc != null`）和 `type:` 伪键：它们不是 Me 的成员。
        /// </summary>
        public Dictionary<string, KeyValuePair<string, int>> Members()
        {
            var map = new Dictionary<string, KeyValuePair<string, int>>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in Decls)
            {
                if (d.Proc != null) continue;
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
        // ★v150：尾部那个 `(?=\s*(?:As|\(|$))` 前瞻必须去掉 —— 它要求**首名**
        //   后面紧跟 As/(/行尾，于是 `Public a, b As Long` 这种逗号多名声明
        //   整行匹配不上，a 和 b 就都没进过候选池（用户报的"公有变量不提示"）。
        //   改成把修饰符之后**整段**交给 DeclTypes 按逗号切，与 Dim 分支同一套口径。
        //   ⚠️ 负面前瞻仍然保留：`Public Sub` / `Private Type` / `Friend Const`
        //   这些不是变量声明，交给后面的过程/Type/Const 分支各自处理。
        private static readonly Regex ReModuleBare = new Regex(
            @"^[ \t]*(Public|Private|Global|Friend)[ \t]+(?!(?:Sub|Function|Property|Declare|Type|Enum|Const|Dim|WithEvents|Event)\b)" +
            @"(.+)$",
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

        // ---- v155：与 PY 版 _RE_IMP_ASSIGN / _RE_IMP_FOR / _RE_IMP_FOR_EACH 等价 ----
        // 只认**写入**（赋值目标、循环变量）。"只读用法"另走宽扫描，靠下面两张名单压噪音。
        private static readonly Regex ReImpAssign = new Regex(
            @"^[ \t]*(?:Let[ \t]+|Set[ \t]+)?(" + Id + @")\s*(?:\([^()]*\))?\s*=(?!=)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ReImpForEach = new Regex(
            @"^[ \t]*For[ \t]+Each[ \t]+(" + Id + @")\s+In\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ReImpFor = new Regex(
            @"^[ \t]*For[ \t]+(" + Id + @")\s*=",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Office / VBA 宿主的**全局对象与集合**名。它们不是用户变量，裸用也不该进
        /// 候选池（`Range("A1")`、`ActiveSheet.UsedRange` 用户天天写，收进池子就是噪音）。
        /// 名单与 PY 版 `_IMPLICIT_STOPWORDS` 里的宿主部分对齐。
        /// </summary>
        private static readonly HashSet<string> HostObjectNames = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            "application", "worksheetfunction", "activeworkbook", "activesheet", "activecell",
            "selection", "range", "cells", "worksheets", "workbooks", "sheets", "charts",
            "chartobjects", "thisworkbook", "rows", "columns", "activewindow", "err", "debug",
            "names", "windows", "shapes", "pages", "menubar", "dialogs", "commandbars",
        };

        private static bool IsHostObjectName(string name) =>
            !string.IsNullOrEmpty(name) && HostObjectNames.Contains(name);

        /// <summary>
        /// 宿主类型库的枚举常量：`vbYes` / `xlUp` / `msoTrue` / `wdAlignParagraphLeft`。
        /// 判据与 PY 版 `_RE_BUILTIN_CONST = ^(?:vb|xl|mso|db|wd|pp|ol|ac|wpp)[A-Z]` 一致：
        /// 前缀全小写、紧跟一个大写字母。用户口径"枚举不要作为提示词"—— 这就是那道闸。
        /// </summary>
        private static readonly string[] HostEnumPrefixes =
            { "vb", "xl", "mso", "db", "wd", "pp", "ol", "ac", "wpp" };

        private static bool IsHostEnumConst(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            foreach (var p in HostEnumPrefixes)
                if (name.Length > p.Length &&
                    name.StartsWith(p, StringComparison.Ordinal) &&
                    char.IsUpper(name[p.Length]))
                    return true;
            return false;
        }

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
                    bool isPriv = bm.Groups[1].Value.Equals("Private", StringComparison.OrdinalIgnoreCase);
                    // 多名声明 `Public a, b As Long` 在这里一起切开（VBA 口径：
                    // 没写 As 的那个是 Variant，DeclTypes 已经这么处理了）。
                    foreach (var v in DeclTypes(bm.Groups[2].Value))
                    {
                        ms.Decls.Add(new Decl { Name = v.Key, Proc = null, IsPrivate = isPriv, TypeName = v.Value });
                        if (v.Value != null) ms.VarTypes["|" + v.Key] = v.Value;
                    }
                    continue;
                }
            }

            ms.OptionExplicit = Parser.OptionExplicit(text);

            // ---- 全量标识符：只留"没被任何过程声明为局部"的 ----
            var declared = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in ms.Decls) declared.Add(d.Name);

            // ---- v153：先分清"这个名字在模块里是怎么用的" ----
            // `Range("d2")`、`InputBox("x")` 这种后面紧跟 `(` 的，是**函数/过程调用**，
            // 不是"自动声明的变量"。无 Option Explicit 时把它们当变量收进池子，
            // 就会拿工程名档位（rank 2）压过词库里的真函数 —— 用户打 `ra`，
            // 第一条是 Excel 的 `Range`，把 `Rate`/`Randomize` 全顶到后面去。
            // 判据只在"这个名字从没被当值用过"时才生效：`arr(1) = 5` 里的 arr
            // 如果也在别处出现 `arr = Array()`，仍然按变量收。
            var callOnly = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var usedAsValue = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var raw in lines)
            {
                string c0 = Parser.CodeOnly(raw);
                if (c0.Length == 0) continue;
                foreach (Match m0 in ReIdent.Matches(c0))
                {
                    string w0 = Clean(m0.Value);
                    if (w0.Length == 0) continue;
                    // `As Range` / `Dim x As Range` 里的 Range 是**类型名**：
                    // 它既不是函数调用、也不是值 —— 归类成"值"会让 `Range("d2")`
                    // 那条判据失效（`!usedAsValue` 永远不成立）。
                    if (IsAfterAs(c0, m0)) continue;
                    // `Implements IFoo`：接口名不是变量
                    if (PrevWordIs(c0, m0, "Implements")) continue;
                    // `GoTo done` / `On Error GoTo fail`：那是**行标签**，不是变量
                    if (PrevWordIs(c0, m0, "GoTo") || PrevWordIs(c0, m0, "GoSub") ||
                        PrevWordIs(c0, m0, "Resume"))
                        continue;
                    // `错误处理:` 这种**标签定义**整行也只有它自己，同样不是变量
                    if (IsLabelDefinition(c0, m0)) continue;

                    if (FollowedByParen(c0, m0))
                    {
                        // ★v154：`arr(1) = 5` 与 `ReDim arr(10)` 里的 arr 是**数组变量**，
                        //   不是函数调用。只看"后面有没有括号"会把它们误判成调用，
                        //   于是"只写下标、从不写 arr = ..."的数组变量一个都提示不出来
                        //   （`Cells(1,1)` 那种宿主函数则相反，本来就该排除）。
                        if (IsArrayLikeUse(c0, m0)) usedAsValue.Add(w0);
                        else callOnly.Add(w0);
                    }
                    else usedAsValue.Add(w0);
                }
            }

            // ---- v155（对齐 PY 的 _implicit_names_in_stmt）：先扫"被写入"的名字 ----
            // 【为什么需要】PY 版的隐式变量分两路：写入（赋值目标 / For / For Each）
            // 是**精确**的，只有"只读用法"才走宽扫描。宽扫描里它靠三张名单压噪音：
            // 宿主对象名（range/cells/…）、枚举常量前缀（vb*/xl*/mso*…）、
            // 和一个"前导词"黑名单（as/new/lib/alias/…）。
            //
            // CS 之前只有一个"全量标识符"，于是 Worksheet 里 `Range("d2")` 和
            // 实参 `xlUp` 都会变成"变量候选"——用户对照 PY 版一眼就看出来了。
            // 这里补上同样的三道闸，并且保留一条例外：只要这个名字在**本模块里
            // 被赋过值**（`range = 1` 这种怪用法），照样按变量收。
            var written = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var raw in lines)
            {
                string cw = Parser.CodeOnly(raw);
                if (cw.Length == 0) continue;
                Match mw = ReImpForEach.Match(cw);
                if (!mw.Success) mw = ReImpFor.Match(cw);
                if (!mw.Success) mw = ReImpAssign.Match(cw);
                if (!mw.Success) continue;
                string nw = Clean(mw.Groups[1].Value);
                if (nw.Length > 0) written.Add(nw);
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool inFreeTypeBlock = false;              // Type 块里的成员不能当裸名（只能 p.X 访问）
            for (int ln = 0; ln < lines.Length; ln++)
            {
                string code = Parser.CodeOnly(lines[ln]);
                if (code.Length == 0) continue;
                // `Attribute VB_Name = "模块1"` 只出现在导出/导入的源码里：
                // 里面的名字不是变量（VBE 的 CodeModule 里看不到这些行）。
                if (code.TrimStart().StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
                    continue;
                // `Option Base 1` / `Option Compare Text`：都是模块设置，里面的词不是变量
                // （`Option Explicit` 由 Parser.OptionExplicit 单独判，跳过这里没有影响）。
                if (code.TrimStart().StartsWith("Option ", StringComparison.OrdinalIgnoreCase))
                    continue;

                // ★v154：`Type Person` 块里的 `Name As String` 是**成员**，不是变量。
                //   这一遍原来完全不看块结构，于是每个用户自定义类型都会往裸名池里
                //   塞进一堆成员名（打 `na` 冒出 `Name`），与 v137 修的"变量泄露"同类。
                //   Enum 块**不**跳过：枚举成员在模块里本来就是裸名（`clRed` 直接用）。
                if (inFreeTypeBlock)
                {
                    if (ReEndType.IsMatch(code)) inFreeTypeBlock = false;
                    continue;
                }
                var tmHead = ReTypeHead.Match(code);
                if (tmHead.Success &&
                    tmHead.Groups[2].Value.Equals("Type", StringComparison.OrdinalIgnoreCase))
                {
                    inFreeTypeBlock = true;
                    continue;                       // Type 名字本身由 Decls 收
                }
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

                    // ★v154：下面这几类**按位置**判断（不能只看名字）：
                    //   同一个 `Range` 既可能是 `As Range` 的类型、也可能是调用的函数，
                    //   所以判据必须落在"这一次出现"上，而不是名字集合上。
                    if (IsAfterAs(code, m)) continue;                        // As Xxx
                    if (PrevWordIs(code, m, "Implements")) continue;         // Implements IFoo
                    if (PrevWordIs(code, m, "GoTo") ||
                        PrevWordIs(code, m, "GoSub") ||
                        PrevWordIs(code, m, "Resume")) continue;             // GoTo 标签
                    if (IsLabelDefinition(code, m)) continue;                // `错误处理:`

                    if (VbaBuiltins.IsKeyword(w)) continue;
                    // ★v154：词库里有的名字（函数 / 类型 / 关键字）不需要再从"全量
                    //   标识符"这条路进来 —— 它们本来就以 builtin 档在池子里。
                    //   收进来的唯一效果是**多一条 var 副本**：用户打 `le` 会同时看到
                    //   `Left`（函数）和 `Left`（变量），去重后保留的是档位更前的那条。
                    //   顺带把 `On Error GoTo …` 里的 `Error` 这类也挡掉。
                    if (VbaBuiltins.IsBuiltin(w)) continue;
                    // `Set x = New Cls`：Cls 是**类名**（组件名），不是变量
                    if (PrevWordIs(code, m, "New")) continue;
                    // `&HFF` / `1E10`：名字其实是数字字面量的一部分
                    if (IsAfterNumber(code, m)) continue;
                    // ★v155：宿主全局对象（Range / Cells / ActiveSheet…）与宿主枚举
                    //   常量（vbYes / xlUp / msoTrue…）不是用户变量。除非本模块里
                    //   给这个名字赋过值（`range = 1`），否则不收 —— 这正是 PY 版
                    //   `_IMPLICIT_STOPWORDS` + `_RE_BUILTIN_CONST` 两道闸的作用。
                    if (!written.Contains(w) && (IsHostObjectName(w) || IsHostEnumConst(w)))
                        continue;
                    if (declared.Contains(w)) continue;
                    if (localNames.Contains(w)) continue;             // 别处的局部，不能漏过来

                    // 单字母（且是 ASCII）噪音太大 —— `i` `n` 这种在候选里没意义。
                    // ★汉字单字不在此列：`和` / `总` 是正经名字，不能按字母数筛掉。
                    if (w.Length < 2 && w[0] < 0x80) continue;

                    bool host = LooksLikeHostObject(code, m);

                    // ★v153：只被当函数调用过、从没当值用过的名字，不是"自动声明的
                    //   变量"（`Range("d2")` / `InputBox("x")`）。收进池子只会让宿主
                    //   函数混进变量档位。真正声明过的过程不受影响 —— 它们走 Decls。
                    if (!host && callOnly.Contains(w) && !usedAsValue.Contains(w)) continue;

                    // ★强制声明变量的模块：没写 Dim 的名字在 VBE 里直接编译不过，
                    //   把它当候选等于教用户写错代码。唯一例外是"像宿主对象"的
                    //   那些（控件名 / Application / 隐式对象）—— 它们本来就不用声明。
                    if (HonorExplicit && ms.OptionExplicit && !host) continue;

                    if (seen.Add(w))
                    {
                        ms.FreeIdents.Add(w);
                        // ★v155（对齐 PY 的 scope="proc"）：未声明变量的作用域是它
                        // **首次出现的过程** —— VBA 里 A 过程的 num 和 B 过程的 num
                        // 是两个独立变量，把 A 的提示到 B 去等于教用户写错代码。
                        ms.FreeIdentProc[w] = ln < owner.Length ? owner[ln] : null;
                    }
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
        /// 这个标识符是不是紧跟在 `As` 后面（= 类型名，不是值也不是调用）。
        /// 只看前一个词是不是 `As`，不做别的推断。
        /// </summary>
        private static bool IsAfterAs(string code, Match m) => PrevWordIs(code, m, "As");

        /// <summary>标识符前面那个词是不是 <paramref name="word"/>（跳过空白）。</summary>
        private static bool PrevWordIs(string code, Match m, string word)
        {
            int p = m.Index - 1;
            while (p >= 0 && (code[p] == ' ' || code[p] == '\t')) p--;
            int e = p;
            while (p >= 0 && (char.IsLetterOrDigit(code[p]) || code[p] == '_')) p--;
            int len = e - p;
            if (len != word.Length) return false;
            return code.Substring(p + 1, len).Equals(word, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 后面跟括号的这个名字，其实是**数组变量**而不是函数调用吗？
        /// 两条判据（都只看语法，不猜类型）：
        ///   * 前面是 `ReDim` / `ReDim Preserve` —— `ReDim arr(10)`；
        ///   * 位于语句起点、且括号配平后紧跟 `=` —— `arr(1) = 5`（数组元素赋值）。
        /// `If arr(1) = 5 Then` 不算（前面有 `If`，不是语句起点）—— 那种情况
        /// 仍归"调用"，但通常同一个 arr 在别处也被赋值过，不影响。
        /// </summary>
        private static bool IsArrayLikeUse(string code, Match m)
        {
            if (PrevWordIs(code, m, "ReDim")) return true;
            if (PrevWordIs(code, m, "Preserve")) return true;      // ReDim Preserve arr(10)
            if (PrevWordIs(code, m, "Let")) return true;           // Let arr(1) = 5
            if (!IsAtStatementStart(code, m)) return false;

            int after = m.Index + m.Length;
            while (after < code.Length && (code[after] == ' ' || code[after] == '\t')) after++;
            if (after >= code.Length || code[after] != '(') return false;

            int depth = 0;
            for (int i = after; i < code.Length; i++)
            {
                if (code[i] == '(') depth++;
                else if (code[i] == ')')
                {
                    if (--depth != 0) continue;
                    int j = i + 1;
                    while (j < code.Length && (code[j] == ' ' || code[j] == '\t')) j++;
                    return j < code.Length && code[j] == '=';
                }
            }
            return false;
        }

        /// <summary>
        /// `名字:` 后面再没有别的内容 —— 这是 VBA 的**行标签定义**（`错误处理:`），
        /// 不是变量。`Select Case` 的 `Case 1:` 之类不受影响（1 不是标识符）。
        /// </summary>
        private static bool IsLabelDefinition(string code, Match m)
        {
            int after = m.Index + m.Length;
            while (after < code.Length && (code[after] == ' ' || code[after] == '\t')) after++;
            return after < code.Length && code[after] == ':' &&
                   code.Substring(after + 1).Trim().Length == 0;
        }

        /// <summary>
        /// 这个名字是不是数字字面量的一部分：`&HFF` 的 `HFF`、`1E10` 的 `E10`。
        /// 只看前一个字符（`&` 或数字），不做进制的额外推断。
        /// </summary>
        private static bool IsAfterNumber(string code, Match m)
        {
            int p = m.Index - 1;
            if (p < 0) return false;
            char c = code[p];
            return char.IsDigit(c) || c == '&';
        }

        /// <summary>这个标识符前面（到行首或冒号为止）只有空白吗 —— 即它在语句起点。</summary>
        private static bool IsAtStatementStart(string code, Match m)
        {
            for (int i = m.Index - 1; i >= 0; i--)
            {
                char c = code[i];
                if (c == ' ' || c == '\t') continue;
                return c == ':';
            }
            return true;
        }

        /// <summary>
        /// 这个标识符后面是不是紧跟一个 `(` —— 也就是它被当成**函数/过程调用**用了。
        /// 只跳过空白，不做任何类型推断（`Range ("d2")` 这种带空格的写法也认）。
        /// 与 <see cref="LooksLikeHostObject"/> 的区别：那边看"后面跟点号"（对象），
        /// 这边看"后面跟括号"（调用）。
        /// </summary>
        private static bool FollowedByParen(string code, Match m)
        {
            int after = m.Index + m.Length;
            while (after < code.Length && (code[after] == ' ' || code[after] == '\t')) after++;
            return after < code.Length && code[after] == '(';
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
