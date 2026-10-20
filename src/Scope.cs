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
        /// 档位给 2：能用，但不如真声明过的名字靠前。</summary>
        public List<string> FreeIdents = new List<string>();
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
        private const string Id = @"[A-Za-z_][A-Za-z0-9_]*[\%\&\#\@\$]?";

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
                    if (w.Length < 2) continue;                       // 单字母噪音太大
                    if (VbaBuiltins.IsKeyword(w)) continue;
                    if (declared.Contains(w)) continue;
                    if (localNames.Contains(w)) continue;             // 别处的局部，不能漏过来
                    if (seen.Add(w)) ms.FreeIdents.Add(w);
                }
            }
            return ms;
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
