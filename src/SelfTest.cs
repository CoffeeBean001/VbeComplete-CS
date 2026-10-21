// 自检：把"补全准不准"的全部判据跑一遍，结果写进 selftest.log。
//
// 用文件而不是 Console：程序是 WinExe（没有控制台），而自检必须在没有桌面的
// 环境下也能跑。

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VbeComplete
{
    internal static class SelfTest
    {
        private static int _pass, _fail;
        private static readonly StringBuilder _sb = new StringBuilder();

        private static void Ok(string name, bool cond, string detail = "")
        {
            if (cond) { _pass++; }
            else
            {
                _fail++;
                _sb.AppendLine("FAIL  " + name + (detail.Length > 0 ? "  <- " + detail : ""));
            }
        }

        private static void Eq<T>(string name, T got, T want)
        {
            bool ok = EqualityComparer<T>.Default.Equals(got, want);
            Ok(name, ok, ok ? "" : "got=" + got + " want=" + want);
        }

        public static int Run()
        {
            _sb.AppendLine("=== VbeComplete C# selftest ===");
            _sb.AppendLine("kernel " + Environment.Version + "  64bit=" + Environment.Is64BitProcess);
            _sb.AppendLine("home   " + Settings.Home);
            _sb.AppendLine();

            try
            {
                TestRules();
                TestParser();
                TestEngine();
                TestFuzzy();
                TestScope();
                TestChineseAndExplicit();
                TestMember();
                TestEditing();
                TestIndent();
                TestAutoTrigger();
                TestSnippets();
            }
            catch (Exception e)
            {
                _fail++;
                _sb.AppendLine("EXCEPTION " + e);
            }

            _sb.AppendLine();
            _sb.AppendLine($"PASS {_pass}   FAIL {_fail}");
            string path = Path.Combine(Settings.Home, "selftest.log");
            try { File.WriteAllText(path, _sb.ToString(), new UTF8Encoding(false)); } catch { }
            return _fail == 0 ? 0 : 1;
        }

        // ------------------------------------------------------------------
        private static void TestRules()
        {
            int ws;
            Eq("word-before/末尾", Rules.WordBefore("Dim abc", 8, out ws), "abc");
            Eq("word-before/起始列", ws, 5);
            Eq("word-before/词中间", Rules.WordBefore("abcdef", 4, out ws), "abc");
            Eq("word-before/无词", Rules.WordBefore("x = ", 5, out ws), "");
            Eq("word-before/空行", Rules.WordBefore("", 1, out ws), "");

            Eq("replace-word", Rules.ReplaceWord("Dim abc", 5, 8, "abcd"), "Dim abcd");
            Eq("replace-word/带后缀", Rules.ReplaceWord("MsgBox \"x\",vbExclamation:Exit Su", 31, 33, "Sub"),
                "MsgBox \"x\",vbExclamation:Exit Sub");

            Eq("dispcol/纯ascii", Rules.DispCol("    abc", 7), 8);
            Eq("dispcol/tab", Rules.DispCol("\tabc", 3), 7);
            Eq("dispcol/宽字符", Rules.DispCol("中文a", 3), 6);

            Eq("find-end/同位置", Rules.FindCompletionEnd("    abc = 1", "abc", 7, 7), 7);
            Eq("find-end/被重排后", Rules.FindCompletionEnd("MsgBox \"已\", vbEx: Exit Sub", "Sub", 24, 22), 26);
            // ★v131b：整词匹配 —— 右边还出现同名子串时，旧实现（±12 窗口 +
            //   LastIndexOf）会取到右边那个，把光标甩到词尾之后。
            Eq("find-end/右边同名整词", Rules.FindCompletionEnd("Split=Split(x)", "Split", 0, 0), 5);
            // ★v131b：`=` 补空格后词起点右移（用户报的 "student=spl" 就是这个）。
            Eq("find-end/等号补空格", Rules.FindCompletionEnd("student = Split()", "Split", 8, 13), 15);
            // 落点列 = `(` 的 1-based 显示列。**不能减 1** —— 减了就落进词中间
            // （`Spli|t`，用户截图那个 bug）。
            Eq("caret-col/括号前", Rules.DispCol("student = Split()", 15), 16);

            // ★v131c：SetCaret 的列语义（"第 N 个字符之前"）就靠这一个函数，
            //   以前少了这个 -1，真机上每个 SetCaret 都稳定右偏一格。
            Eq("caret-col/第1个字符之前", Rules.CaretColForCharCol("    n = 1", 1), 1);
            Eq("caret-col/第5个字符之前", Rules.CaretColForCharCol("    n = 1", 5), 5);
            Eq("caret-col/行尾之后", Rules.CaretColForCharCol("    Dim abcd", 13), 13);
            Eq("caret-col/全角行", Rules.CaretColForCharCol("    学生 = 1", 7), 9);

            // ★v131b：VBE 给的是显示列。纯 ASCII 行两种语义数值相同（所以这个坑
            //   只在有中文/全角的行上暴露）；全角行必须换算，否则取词会落空。
            Eq("char-col/纯ascii 恒等", Rules.CharColFromDisp("student = Split()", 16), 16);
            Eq("char-col/全角缩进", Rules.CharColFromDisp("    学生 = Split()", 17), 15);
            Eq("char-col/行尾", Rules.CharColFromDisp("abc", 4), 4);
            Eq("char-col/超出钳到行尾", Rules.CharColFromDisp("abc", 99), 4);
            // 换算之后再取词 —— 这一对是"含中文的行到底弹不弹"的判据
            Eq("char-col/全角行取词", Rules.WordBefore("    学生 = Split()",
                Rules.CharColFromDisp("    学生 = Split()", 17), out _), "Split");

            FuzzyHit mh;
            Ok("match/前缀", Fuzzy.Match("MsgBox", "msg", out mh) && mh.First == 0 && mh.Kind == 3);
            Ok("match/边界跳步", Fuzzy.Match("selBegin", "sb", out mh) &&
                                 mh.First == 0 && mh.Span == 4 && mh.Kind == 2);
            Ok("match/非前缀单字不算前缀", Fuzzy.Match("abcDef", "b", out mh) &&
                                           mh.First == 1 && mh.Kind < 3);
            Ok("match/不命中", !Fuzzy.Match("abc", "xyz", out mh));

            Ok("resolved/精确且无更长",
                Rules.PrefixIsResolved("sub", new List<string> { "sub", "function" }));
            Ok("resolved/有更长候选不算打全",
                !Rules.PrefixIsResolved("sub", new List<string> { "sub", "subset" }));
            Ok("resolved/没精确命中",
                !Rules.PrefixIsResolved("su", new List<string> { "sub" }));

            // ★v138：成员位置判据（点号后整个让给 VBE）—— 从 Python 版
            //   `_member_list_dot_col` 移植，认 / 不认的位置各钉几条。
            Ok("member-position/点号后", Rules.InMemberPosition("student.", 9));
            // ★改口径：「词中间」也是成员位置（VBE 这时正在按已打的字过滤成员），
            //   旧判据只认"紧跟点号"，于是 `student.na|` 会弹出我们的裸名池。
            Ok("member-position/成员名打到一半", Rules.InMemberPosition("student.na", 10));
            Ok("member-position/空格后点号", Rules.InMemberPosition("a = b. ", 8));
            Ok("member-position/点号带空格", Rules.InMemberPosition("a.b .c", 6));
            Ok("member-position/) 之后", Rules.InMemberPosition("    Range(\"A1\").", 19));
            Ok("member-position/] 之后", Rules.InMemberPosition("    arr(0).Cou", 15));
            // With 块：点顶在行首 / 紧跟语句边界 —— VBE 同样会弹成员列表
            Ok("member-position/With 块行首点", Rules.InMemberPosition("    .Size", 10));
            Ok("member-position/With 块赋值点", Rules.InMemberPosition("    Set f = .Font", 17));
            Ok("member-position/语句分隔后的点", Rules.InMemberPosition("    a = 1: .Size", 15));
            // 反例：这些点号不是成员访问，必须照常提示
            Ok("member-position/小数点不是", !Rules.InMemberPosition("x = 1.5", 8));
            Ok("member-position/没有点号不是", !Rules.InMemberPosition("student", 8));
            Ok("member-position/运算符后不是", !Rules.InMemberPosition("x = a + b", 10));
            Ok("member-position/串字面量后不是", !Rules.InMemberPosition("x = \"abc\".c", 13));
            Ok("member-position/注释里的点不是", !Rules.InMemberPosition("x = 1 ' c.d", 12));
            Ok("member-position/字符串里的点不是", !Rules.InMemberPosition("s = \"a.b\"", 10));
            Ok("member-position/返回点号下标", Rules.MemberDotCol("student.na", 10) == 7);

            Ok("decl-position/Dim 名字", Rules.InDeclarationPosition("Dim cus", 8));
            // ★v133：`As` 后面一律不弹（用户明确要求）—— 那里只可能是类型名
            Ok("decl-position/As 后不弹", Rules.InDeclarationPosition("Dim x As Inte", 13));
            Ok("decl-position/As 后空词也不弹", Rules.InDeclarationPosition("Dim x As ", 11));
            Ok("decl-position/Sub 名字", Rules.InDeclarationPosition("Sub Do", 7));
            Ok("decl-position/普通语句不是", !Rules.InDeclarationPosition("x = cus", 8));
            Ok("decl-position/Const 值位置", !Rules.InDeclarationPosition("Const PI = 3", 13));
            // ★v133：不带 Dim 的裸声明 —— 旧正则漏了这几条，于是 `Public glo|`
            //   被当成普通表达式，弹窗把"正在定义的那个名字自己"提示出来。
            Ok("decl-position/裸 Public", Rules.InDeclarationPosition("Public gCount", 13));
            Ok("decl-position/裸 Private", Rules.InDeclarationPosition("Private m_arr", 13));
            Ok("decl-position/裸 Global", Rules.InDeclarationPosition("Global g", 8));
            Ok("decl-position/裸 Public As", Rules.InDeclarationPosition("Public g As Long", 16));
            Ok("decl-position/Public Sub 不算裸声明误判", Rules.InDeclarationPosition("Public Sub DoIt(", 16));

            // ★v137：`As` / `New` 之后一律不弹 —— 与 decl-position 互补：
            //   那条认"声明语句的行首"，这条认 `Set c = New |` / `For i As |` /
            //   过程参数 `a As |` 这些行首不是声明关键字的位置（用户报的第 2 条）。
            Ok("type-kw/New 后（空词）", Rules.AfterTypeKeyword("    Set c = New ", 17));
            Ok("type-kw/New 后半截词", Rules.AfterTypeKeyword("    Set c = New cl", 17));
            Ok("type-kw/As 后（For 行）", Rules.AfterTypeKeyword("    For i As Lo", 14));
            Ok("type-kw/过程参数 As 后", Rules.AfterTypeKeyword("Sub f(a As Lo", 12));
            Ok("type-kw/New 大小写", Rules.AfterTypeKeyword("x = NEW cl", 9));
            // 反例：这几个都**不是**关键字位置，必须照常提示
            Ok("type-kw/空行不是", !Rules.AfterTypeKeyword("", 1));
            Ok("type-kw/普通语句不是", !Rules.AfterTypeKeyword("x = newValue", 5));
            Ok("type-kw/Renew 不算 New", !Rules.AfterTypeKeyword("x = Renew", 5));
            Ok("type-kw/As 出现在词中间不算", !Rules.AfterTypeKeyword("x = ClassName", 5));

            // 把"正在输入的词"挖掉（自我提示的根治办法）
            Eq("blank-token/挖掉当前词",
                Rules.BlankToken("Public glo\r\nx = 1", 1, 8, "glo"), "Public    \nx = 1");
            Eq("blank-token/位置对不上就不动",
                Rules.BlankToken("Public abc\r\nx = 1", 1, 8, "glo"), "Public abc\r\nx = 1");
            Eq("blank-token/中文词",
                Rules.BlankToken("    学生 = 1", 1, 5, "学生"), "      " + " = 1");

            Ok("comment/注释里", Rules.InCommentOrString("' abc", 5));
            Ok("comment/字符串里", Rules.InCommentOrString("x = \"ab\"", 8));
            Ok("comment/普通代码不是", !Rules.InCommentOrString("x = 1", 5));

            char ch;
            Ok("line-edit/插入一个字符", Rules.LineEditInserted("Dim ab", "Dim abc", out ch) && ch == 'c');
            Ok("line-edit/删除不算", !Rules.LineEditInserted("Dim abc", "Dim ab", out ch));
            Ok("line-edit/替换不算", !Rules.LineEditInserted("Dim ab", "Dim xb", out ch));

            Ok("call-site/赋值右侧", Rules.IsCallSite("x = getVa", 9));
            Ok("call-site/点号后不是", !Rules.IsCallSite("a.b", 3));
            Ok("call-site/声明位置不是", !Rules.IsCallSite("Sub foo", 5));
            Ok("call-site/As 之后不是", !Rules.IsCallSite("Dim x As Long", 10));
            // 行首（只有缩进）也是调用点 —— "函数提示词不补括号"就栽在这条上
            Ok("call-site/行首是", Rules.IsCallSite("    MsgB", 5));
            Ok("call-site/顶格行首是", Rules.IsCallSite("MsgB", 1));
            Ok("call-site/注释之后不是", !Rules.IsCallSite("Call MsgB", 6));
        }

        // ------------------------------------------------------------------
        private const string Sample = @"Option Explicit
Private mCount As Long

Public Sub Main()
    Dim a As Long, b As String
    a = 1
    Call Helper(a)
    Set obj = New Collection
End Sub

Private Function GetName(ByVal idx As Long, Optional ByVal tag As String = """") As String
    GetName = ""x""
End Function

Public Property Get Size() As Long
    Size = 1
End Property

Public Sub Main()
    ' 同名过程的第二个
End Sub

Public Enum Color2
    clRed = 1
    clBlue
End Enum

Public Type Point2
    X As Long
    Y As Long
End Type
";

        private static void TestParser()
        {
            var procs = Parser.ExtractProcedures(Sample);
            Eq("procs/数量", procs.Count, 4);
            Eq("procs/第一个名字", procs[0].Name, "Main");
            Eq("procs/第一个种类", procs[0].Kind, "sub");
            Eq("procs/函数种类", procs[1].Kind, "function");
            Eq("procs/属性种类", procs[2].Kind, "property");
            Ok("procs/同名第二个带序号", procs[3].ScopeKey == "Main#2", procs[3].ScopeKey);

            var p = Parser.ProcOfLine(procs, 6);
            Ok("procs/第 6 行在 Main 里", p != null && p.Name == "Main", p?.ToString());

            var decl = Parser.ExtractDeclared(Sample);
            Ok("decl/多声明 a", decl.Contains("a"));
            Ok("decl/多声明 b", decl.Contains("b"));
            Ok("decl/参数 idx", decl.Contains("idx"));
            Ok("decl/参数 tag", decl.Contains("tag"));
            Ok("decl/枚举成员", decl.Contains("clRed"));
            Ok("decl/类型成员", decl.Contains("X"));
            Ok("decl/不能有 As", !decl.Contains("As"));
            Ok("decl/不能有 Optional", !decl.Contains("Optional"));

            var ids = Parser.ExtractIdentifiers(Sample);
            Ok("ids/收过程名", ids.Contains("Helper"));
            Ok("ids/收变量", ids.Contains("mCount"));
            Ok("ids/不收关键字", !ids.Contains("Dim"));
            Ok("ids/不收字符串内容", !ids.Contains("同名过程的第二个"));
            Ok("ids/不收注释词", !ids.Contains("同名过程"));

            var types = Parser.ExtractTypeNames(Sample);
            Ok("types/Enum 名", types.Contains("Color2"));
            Ok("types/Type 名", types.Contains("Point2"));

            Ok("option-explicit/有", Parser.OptionExplicit(Sample));
            Ok("option-explicit/没有", !Parser.OptionExplicit("Sub a()\r\nEnd Sub"));

            Eq("code-only/去注释", Parser.CodeOnly("x = 1 ' abc").Trim(), "x = 1");
            Eq("code-only/去字符串", Parser.CodeOnly("x = \"abc\"").Trim(), "x =");
            Eq("code-only/Rem", Parser.CodeOnly("Rem abc"), "");
            Eq("code-only/不误伤 Remark", Parser.CodeOnly("Remark = 1").Trim(), "Remark = 1");
        }

        // ------------------------------------------------------------------
        private static void TestEngine()
        {
            var eng = new Engine
            {
                Recent = new Recent(),
            };
            eng.PrimePool(Sample, "Module1");

            bool wp, ws;
            eng.Followup("Main", "Other", out wp, out ws);
            Ok("followup/sub 补空格", !wp && ws, $"parens={wp} space={ws}");

            eng.Followup("GetName", "Other", out wp, out ws);
            Ok("followup/function 补括号", wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("GetName", "GetName", out wp, out ws);
            Ok("followup/函数体内给自己赋值不补", !wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("GetName", "getname#2", out wp, out ws);
            Ok("followup/作用域键带序号也认得出自己", !wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("Size", "Other", out wp, out ws);
            Ok("followup/property 都不补", !wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("MsgBox", "Other", out wp, out ws);
            Ok("followup/内建函数补括号", wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("String", "Other", out wp, out ws);
            Ok("followup/双身份名字不补括号", !wp, $"parens={wp}");

            eng.Followup("Public", "Other", out wp, out ws);
            Ok("followup/关键字补空格", !wp && ws, $"parens={wp} space={ws}");

            eng.Followup("Next", "Other", out wp, out ws);
            Ok("followup/Next 不补空格", !wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("Me", "Other", out wp, out ws);
            Ok("followup/Me 不补空格", !wp && !ws, $"parens={wp} space={ws}");

            Eq("classify/工程过程", eng.Classify("Main"), "func");
            Eq("classify/工程属性", eng.Classify("Size"), "prop");
            Eq("classify/工程类型", eng.Classify("Color2"), "type");
            Eq("classify/内建函数", eng.Classify("MsgBox"), "func");
            Eq("classify/内建类型", eng.Classify("Long"), "type");
            Eq("classify/关键字", eng.Classify("Public"), "kw");
            Eq("classify/变量", eng.Classify("mCount"), "var");

            // 排序：工程内的名字必须排在内建前面
            var a = new Candidate { Name = "Main", ScopeKind = 3, Pos = 0, Span = 2, Kind = "func" };
            var b = new Candidate { Name = "MsgBox", ScopeKind = 2, Pos = 0, Span = 2, Kind = "func" };
            Ok("sort/工程名优先于内建", Engine.Cmp(a, b) < 0);

            var c1 = new Candidate { Name = "abc", ScopeKind = 3, Pos = 0, Span = 1 };
            var c2 = new Candidate { Name = "abc", ScopeKind = 3, Pos = 1, Span = 1 };
            Ok("sort/命中位置靠前优先", Engine.Cmp(c1, c2) < 0);

            var r1 = new Candidate { Name = "a", ScopeKind = 3, Pos = 0, Span = 1, RecentRank = 5 };
            var r2 = new Candidate { Name = "b", ScopeKind = 3, Pos = 0, Span = 1, RecentRank = 1 };
            Ok("sort/最近使用优先", Engine.Cmp(r1, r2) < 0);

            // ★命中质量必须排在作用域**前面**：打 `msgbox` 时，一个工程内的
            //   长名字只是"松散命中"，绝不能因为它"工程内"就顶掉精确命中的 MsgBox。
            var loose = new Candidate { Name = "MsgBoxLocalHelper", ScopeKind = 4, HitKind = 0, Pos = 3, Span = 9 };
            var exact = new Candidate { Name = "MsgBox", ScopeKind = 2, HitKind = 3, Pos = 0, Span = 6 };
            Ok("sort/命中质量优先于作用域", Engine.Cmp(exact, loose) < 0);

            // 名字短的在前（旧代码写反了：长度降序 = 长的在前）
            var shortN = new Candidate { Name = "ab", ScopeKind = 3, HitKind = 3, Pos = 0, Span = 2 };
            var longN = new Candidate { Name = "abcdef", ScopeKind = 3, HitKind = 3, Pos = 0, Span = 8 };
            Ok("sort/名字短的在前", Engine.Cmp(shortN, longN) < 0);

            // ---- 启动脚本契约（v131a）----
            // "双击窗口一闪、项目没启动"出过两次，两次都在 .bat 上：
            //   v129：echo 文本里一个裸右括号就让 cmd 丢弃整个 if 块并退出 255；
            //   v131a：LF-only 换行让 goto 标签和多行 if 块失效；
            //   还有更根本的一条：start 之后既不验证也不提示，程序真没起来用户也看不到。
            // 这些是【文件形态】问题，只能查字节，别用源码字符串护栏（注释会喂饱它）。
            var root = Settings.Home;
            var bats = Directory.GetFiles(root, "*.bat", SearchOption.AllDirectories)
                .Where(p => !p.Contains(@"\obj\") && !p.Contains(@"\bin\"))
                .ToArray();
            Ok("bat/找到启动脚本", bats.Length >= 4, bats.Length + " 个");

            var badEol = new List<string>();
            var badParen = new List<string>();
            foreach (var f in bats)
            {
                var bytes = File.ReadAllBytes(f);
                int lfOnly = 0;
                for (int i = 0; i < bytes.Length; i++)
                {
                    if (bytes[i] != (byte)'\n') continue;
                    if (i == 0 || bytes[i - 1] != (byte)'\r') lfOnly++;
                }
                if (lfOnly > 0) badEol.Add(Path.GetFileName(f) + "(LF=" + lfOnly + ")");

                foreach (var line in File.ReadAllLines(f))
                {
                    var t = line.Trim();
                    if (!t.StartsWith("echo", StringComparison.OrdinalIgnoreCase)) continue;
                    if (t.IndexOf(')') >= 0) badParen.Add(Path.GetFileName(f) + ": " + t);
                }
            }
            Ok("bat/必须是 CRLF 换行", badEol.Count == 0, string.Join(", ", badEol));
            Ok("bat/echo 不许有裸右括号", badParen.Count == 0, string.Join(" | ", badParen));

            var runCs = Path.Combine(root, "run-cs.bat");
            var runCsText = File.Exists(runCs) ? File.ReadAllText(runCs) : "";
            Ok("bat/启动后验证进程存活", runCsText.Contains("goto died") && runCsText.Contains("tasklist"),
                "缺存活校验");
            Ok("bat/失败时打印日志尾部",
                runCsText.Contains("crash.log") && runCsText.Contains("VbeComplete.log"),
                "没有把日志给用户看");
            Ok("bat/不用 find（Git 的同名程序会抢 /i）",
                runCsText.Contains("findstr") && !runCsText.Contains("| find /i"),
                "still uses find");
        }

        // ------------------------------------------------------------------
        // 编辑手感（块结构 / 缩进 / 自动配对 / 退格接管）
        // ------------------------------------------------------------------
        // 模糊匹配：这一层的每个断言都对应 Python 版里一条真实修过的 bug，
        // 改算法前先看一眼 —— 里面有一条是"贪心会漏判"的经典反例。
        // ------------------------------------------------------------------
        private static void TestFuzzy()
        {
            FuzzyHit h;

            // 1) 整块连续优先：arr 命中 dataArr 的 Arr，而不是散着的 a..r..r
            Ok("fuzzy/连续块", Fuzzy.Match("dataArr", "arr", out h));
            Eq("fuzzy/连续块-kind", h.Kind, 1);
            Eq("fuzzy/连续块-位置", string.Join(",", h.Positions), "4,5,6");

            // 2) 词首缩略：ds -> dataSheet 的 D 与 S（中间没有连续的 ds）
            Ok("fuzzy/词首缩略", Fuzzy.Match("dataSheet", "ds", out h));
            Eq("fuzzy/词首缩略-kind", h.Kind, 2);
            Eq("fuzzy/词首缩略-位置", string.Join(",", h.Positions), "0,4");

            // 3) ts -> dataSheet：t 在词中、S 在词边界
            Ok("fuzzy/边界跳步", Fuzzy.Match("dataSheet", "ts", out h));
            Eq("fuzzy/边界跳步-位置", string.Join(",", h.Positions), "2,4");

            // 4) 前缀 / 完全相等
            Ok("fuzzy/前缀", Fuzzy.Match("dataSheet", "dat", out h));
            Eq("fuzzy/前缀-kind", h.Kind, 3);
            Ok("fuzzy/完全相等", Fuzzy.Match("dataSheet", "datasheet", out h));
            Eq("fuzzy/完全相等-kind", h.Kind, 4);
            Ok("fuzzy/大小写无关", Fuzzy.Match("dataSheet", "DATAS", out h));

            // 5) ★贪心漏判的经典反例（Python 版 v59）：
            //    第 3 个字符 t 若贪心跳到词边界 T(10)，后面就没有 s 了；
            //    正确解是 [0,1,2,3,4,15]。反向可行性上界就是为这条加的。
            Ok("fuzzy/子序列不许漏判", Fuzzy.Match("getSettingTitle3", "getse3", out h));
            Eq("fuzzy/子序列-位置", string.Join(",", h.Positions), "0,1,2,3,4,15");
            Eq("fuzzy/子序列-最长连续段", h.Run, 5);

            // 6) 跳步但"真的打过一段"：xlworkfaul -> xlWorkbookDefault
            Ok("fuzzy/长名字跳步", Fuzzy.Match("xlWorkbookDefault", "xlworkfaul", out h));
            Eq("fuzzy/长名字-最长连续段", h.Run, 6);

            // 7) 不匹配 / 长度不够 / 空 query
            Ok("fuzzy/不匹配", !Fuzzy.Match("dataSheet", "zzz", out h));
            Ok("fuzzy/比名字还长", !Fuzzy.Match("abc", "abcd", out h));
            Ok("fuzzy/空 query 全都要", Fuzzy.Match("abc", "", out h));
            Eq("fuzzy/空 query-位置个数", h.Positions.Length, 0);

            // 8) 中文名字：没有大小写/驼峰，整块连续照样命中
            Ok("fuzzy/中文前缀", Fuzzy.Match("学生列表", "学生", out h));
            Eq("fuzzy/中文-kind", h.Kind, 3);

            // 9) 下划线也是词边界（VBA 里 m_Name 很常见）
            Ok("fuzzy/下划线边界", Fuzzy.Match("m_studentName", "sn", out h));
            Eq("fuzzy/下划线边界-位置", string.Join(",", h.Positions), "2,9");

            // 10) 高亮切段：相邻同类必须合并成一段（否则一个名字要画十几个文本项）
            var segs = PopupForm.SplitByHits("dataSheet", new[] { 0, 4 });
            Eq("hits/段数", segs.Count, 4);
            Eq("hits/首段文字", segs[0].Key, "d");
            Ok("hits/首段命中", segs[0].Value);
            Eq("hits/次段文字", segs[1].Key, "ata");
            Ok("hits/次段未命中", !segs[1].Value);
            Eq("hits/M", segs[2].Key, "S");
            Eq("hits/尾段", segs[3].Key, "heet");
            Eq("hits/空名字", PopupForm.SplitByHits("", new[] { 0 }).Count, 0);
        }

        // ------------------------------------------------------------------
        // 作用域：哪些名字在当前位置**看得见**。这里每一条都是"提示里混进
        // 用不了的名字"的真 bug 的反面。
        // ------------------------------------------------------------------
        private const string ScopedSample = @"Option Explicit
Private mModuleLevel As Long
Public gPublic As Long

Private Sub A()
    Dim tmpA As Long
    Dim shared As String
    tmpA = 1
End Sub

Public Sub B(ByVal p1 As Long)
    Dim tmpB As Long
    shared = ""x""
End Sub

Public Sub C()
End Sub
";

        private static void TestScope()
        {
            var ms = Scope.Analyze("Mod1", ScopedSample);

            var va = ms.Visible("A");
            Ok("scope/A 看得见自己的局部", va.ContainsKey("tmpA"));
            Eq("scope/局部档位=4", va["tmpA"].Value, 4);
            Ok("scope/A 看得见模块级 Private", va.ContainsKey("mModuleLevel"));
            Eq("scope/模块级档位=3", va["mModuleLevel"].Value, 3);
            Ok("scope/A 看不见 B 的局部", !va.ContainsKey("tmpB"));
            Ok("scope/A 看得见本模块过程名", va.ContainsKey("B"));

            var vb = ms.Visible("B");
            Ok("scope/参数是局部", vb.ContainsKey("p1"));
            Eq("scope/参数档位=4", vb["p1"].Value, 4);
            Ok("scope/B 看不见 A 的局部", !vb.ContainsKey("tmpA"));
            // ★这条是核心：B 里**引用**了 shared，但 shared 声明在 A 里
            //   —— 不能因为"B 里出现过这个词"就把 A 的局部漏给 B。
            Ok("scope/别处引用过的局部不漏过来", !vb.ContainsKey("shared"));

            var ex = ms.Exported();
            Ok("scope/Public 过程可导出", ex.ContainsKey("B") && ex.ContainsKey("C"));
            Ok("scope/Private 过程不导出", !ex.ContainsKey("A"));
            Ok("scope/Public 变量可导出", ex.ContainsKey("gPublic"));
            Ok("scope/Private 变量不导出", !ex.ContainsKey("mModuleLevel"));
            Ok("scope/局部一律不导出", !ex.ContainsKey("tmpA") && !ex.ContainsKey("p1"));

            // 引擎层面：池子里有 shared，但在 B 的位置不该出现在候选里
            var eng = new Engine { Recent = new Recent() };
            eng.PrimePool(ScopedSample, "Mod1");
            var items = enginesTrigger(eng, "    sha", 8);
            Ok("scope/引擎在 B 里不提示 A 的局部",
                items.All(x => x.Name.ToLowerInvariant() != "shared"),
                string.Join(",", items.Select(x => x.Name)));

            // 同一个词在它自己的过程里必须提示得到
            var items2 = enginesTrigger(eng, "    sha", 8, "A");
            Ok("scope/引擎在 A 里提示得到自己的局部",
                items2.Any(x => x.Name.ToLowerInvariant() == "shared"),
                string.Join(",", items2.Select(x => x.Name)));
        }

        /// <summary>造一个假快照跑 Trigger（不碰 COM），返回候选名。</summary>
        private static List<Candidate> enginesTrigger(Engine eng, string line, int caretDisp,
            string proc = "B")
        {
            int ln = proc == "A" ? 7 : 13;          // 7 在 Sub A 里，13 在 Sub B 里
            var be = new FakeBackend
            {
                Text = ScopedSample,
                Snap = new Snapshot { Line = ln, Text = line, Module = "Mod1", CaretCol = caretDisp },
            };
            eng.Backend = be;
            if (!eng.Trigger(false, be.Snap)) return new List<Candidate>();
            return eng.Items.ToList();
        }

        /// <summary>只给引擎喂快照的最小后端：不碰 COM。</summary>
        private sealed class FakeBackend : IVbeSource
        {
            public string Text;
            public Snapshot Snap;
            public List<ModuleScope> Extra = new List<ModuleScope>();
            public Snapshot Snapshot(bool verify = true) => Snap;
            public string ModuleText() => Text;
            public List<string> ComponentNames() => new List<string>();
            public List<ModuleScope> ProjectScopes(out int signature)
            {
                signature = 0;
                return Extra;
            }
            public bool ApplyCompletion(int lineNo, int wordStartCol, int endCol,
                string completion, bool wantParens, bool wantSpace, string enclosingProc) => true;
            public string EnclosingProcName() => null;
        }

        // ------------------------------------------------------------------
        // `.` 之后的成员：只在"类型写得明明白白"时才补，否则整个让给 VBE。
        // ------------------------------------------------------------------
        private const string MemberSample = @"Option Explicit
Private Sub Main()
    Dim c As Cls1
    Set c = New Cls1
    c.
End Sub

Public Sub DoIt()
End Sub
";

        private const string Cls1Sample = @"Public Sub Refresh()
End Sub

Private Sub Hidden()
End Sub

Public Name As String
";

        // ------------------------------------------------------------------
        // 中文标识符 + Option Explicit（v133）
        // ------------------------------------------------------------------
        private const string ChineseSample = @"Option Explicit
Private 模块计数 As Long

Public Sub 汇总(ByVal 起始行 As Long)
    Dim 合计 As Long
    合计 = 1
End Sub

Public Function 取总分() As Long
    取总分 = 0
End Function
";

        private const string NoExplicitSample = @"Private Sub 测试()
    学生 = 1
    总分 = 学生 + 2
End Sub
";

        private const string HostObjSample = @"Option Explicit
Private Sub 点击()
    CommandButton1.Caption = ""确定""
    Set 按钮 = Me.Controls.Add(""Forms.CommandButton.1"")
End Sub
";

        private static void TestChineseAndExplicit()
        {
            // ---- 中文过程名 / 变量名必须解析出来 ----
            var ms = Scope.Analyze("模块1", ChineseSample);
            Ok("cn/中文过程名是声明", ms.Decls.Any(d => d.Name == "汇总" && d.KindHint == "sub"),
                string.Join(",", ms.Decls.Select(d => d.Name)));
            Ok("cn/中文函数名", ms.Decls.Any(d => d.Name == "取总分" && d.KindHint == "function"));
            Ok("cn/中文模块变量", ms.Decls.Any(d => d.Name == "模块计数" && d.Proc == null));
            Ok("cn/中文参数是局部", ms.Decls.Any(d => d.Name == "起始行" && d.Proc != null));

            var vis = ms.Visible("汇总");
            Ok("cn/中文局部可见", vis.ContainsKey("合计"));
            Ok("cn/中文参数可见", vis.ContainsKey("起始行"));

            // 中文名要能进候选池并被模糊命中
            var cnSnap = new Snapshot
            {
                // ★CaretCol 是**显示列**：4 个空格 + 一个汉字 = 宽 6，光标落在 7
                Line = 6, Text = "    合", Module = "模块1", CaretCol = 7,
            };
            var cnBe = new FakeBackend { Text = ChineseSample, Snap = cnSnap };
            var eng = new Engine { Recent = new Recent(), Backend = cnBe };
            var items = eng.Trigger(false, cnSnap) ? eng.Items.ToList() : new List<Candidate>();
            Ok("cn/打中文首字能提示", items.Any(c => c.Name == "合计"),
                string.Join(",", items.Select(c => c.Name)));

            // ---- Option Explicit ----
            Ok("explicit/认出来了", Scope.Analyze("M", ChineseSample).OptionExplicit);
            Ok("explicit/没写就不算", !Scope.Analyze("M", NoExplicitSample).OptionExplicit);

            var noExp = Scope.Analyze("M", NoExplicitSample);
            Ok("explicit/无该选项时未声明名字进池",
                noExp.FreeIdents.Contains("学生") && noExp.FreeIdents.Contains("总分"),
                string.Join(",", noExp.FreeIdents));

            var expSample = "Option Explicit\r\nPrivate Sub T()\r\n    未定义 = 1\r\nEnd Sub\r\n";
            var exp = Scope.Analyze("M", expSample);
            Ok("explicit/有该选项时未声明名字不进池", !exp.FreeIdents.Contains("未定义"),
                string.Join(",", exp.FreeIdents));

            // ---- 自我提示：正在打的这个词本身不该进候选池 ----
            // ★旧版症状：定义全局/模块变量时（`Public glo|`），解析把半截名当成了
            //   已声明的名字，于是候选第一条就是"你正在打的这个词"，一弹一收像闪烁。
            const string selfSample = "Sub T()\r\n    myNewVar = 1\r\nEnd Sub\r\n";
            var selfBe = new FakeBackend
            {
                Text = selfSample,
                Snap = new Snapshot { Line = 2, Text = "    myNewVa", Module = "M", CaretCol = 13 },
            };
            var selfEng = new Engine { Recent = new Recent(), Backend = selfBe };
            bool selfShown = selfEng.Trigger(false, selfBe.Snap);
            Ok("self/正在输入的名字不提示自己",
                !selfShown || !selfEng.Items.Any(c => c.Name == "myNewVar"),
                selfShown ? string.Join(",", selfEng.Items.Select(c => c.Name)) : "(不弹)");

            // 宿主对象（点了点号的 / Set 的）在 Option Explicit 下仍要留
            var host = Scope.Analyze("M", HostObjSample);
            Ok("explicit/控件名（带点号）保留", host.FreeIdents.Contains("CommandButton1"),
                string.Join(",", host.FreeIdents));
            Ok("explicit/Set 的对象名保留", host.FreeIdents.Contains("按钮"),
                string.Join(",", host.FreeIdents));

            // ---- 引擎层面：Option Explicit 的模块里，未声明名字不该出现在候选 ----
            var eng2 = new Engine { Recent = new Recent() };
            var be = new FakeBackend
            {
                Text = expSample,
                Snap = new Snapshot { Line = 3, Text = "    未定", Module = "M", CaretCol = 7 },
            };
            eng2.Backend = be;
            bool shown = eng2.Trigger(false, be.Snap);
            Ok("explicit/候选里没有未声明名字",
                !shown || !eng2.Items.Any(c => c.Name == "未定义" || c.Name == "未定"),
                shown ? string.Join(",", eng2.Items.Select(c => c.Name)) : "(不弹)");
        }

        private static void TestMember()
        {
            var cls = Scope.Analyze("Cls1", Cls1Sample);
            Eq("member/类型解析", cls.TypeOf("c", null), null);      // 模块级没有 c
            var mod = Scope.Analyze("Mod1", MemberSample);
            Eq("member/Dim As 收类型", mod.TypeOf("c", "Main"), "Cls1");

            // ★v138：默认口径 —— 点号之后**一个候选都不出**，整个让给 VBE 的
            //   成员列表（与 Python 版 v103c 同一口径）。这条必须先钉住：下面
            //   那些"收窄"断言都跑在 MEMBER_POPUP=1 下，不先钉默认口径的话，
            //   一旦判据写错（把 `c.` 判成非成员位置），整套测试都不会红。
            var def = new Engine { Recent = new Recent() };
            var beDef = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    c.", Module = "Mod1", CaretCol = 7 },
            };
            def.Backend = beDef;
            Ok("member/默认点号后不弹（即使类型已知）", !def.Trigger(false, beDef.Snap));
            var beDef2 = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    c.Ref", Module = "Mod1", CaretCol = 10 },
            };
            def.Backend = beDef2;
            Ok("member/默认点号后打了半截词也不弹", !def.Trigger(false, beDef2.Snap));

            // ---- 以下都是 MEMBER_POPUP=1 那条路（用户显式要候选）----
            var eng = new Engine { Recent = new Recent(), MemberPopup = true };
            var be = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    c.", Module = "Mod1", CaretCol = 7 },
            };
            eng.Backend = be;
            bool shown = eng.Trigger(false, be.Snap);
            Ok("member/能确定类型就出成员", shown, "trigger=false");
            var names = eng.Items.Select(x => x.Name).ToList();
            Ok("member/含 Public 过程", names.Contains("Refresh"), string.Join(",", names));
            Ok("member/含 Public 变量", names.Contains("Name"), string.Join(",", names));
            Ok("member/不含 Private 成员", !names.Contains("Hidden"), string.Join(",", names));
            Ok("member/不含全工程噪音", !names.Contains("MsgBox"), string.Join(",", names));

            // 类型不确定（变量从没声明过）=> 退回全工程候选（与 Python 版同一口径：
            // 定不出类型时宁可给宽的，也不给"收窄成空"）
            eng.Hide();
            var be2 = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    obj.", Module = "Mod1", CaretCol = 9 },
            };
            eng.Backend = be2;
            bool shown2 = eng.Trigger(false, be2.Snap);
            Ok("member/类型不确定时退回全工程候选",
                shown2 && eng.Items.Any(x => x.Name == "MsgBox"),
                shown2 ? string.Join(",", eng.Items.Take(8).Select(x => x.Name)) : "(不弹)");

            // 点号后面已经开始打词：收窄仍然生效
            eng.Hide();
            var be3 = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    c.Ref", Module = "Mod1", CaretCol = 10 },
            };
            eng.Backend = be3;
            Ok("member/点号后继续打字仍收窄", eng.Trigger(false, be3.Snap));
            Ok("member/收窄后只剩一个",
                eng.Items.Count == 1 && eng.Items[0].Name == "Refresh",
                string.Join(",", eng.Items.Select(x => x.Name)));

            // `Me.` = 当前模块自己的 Public 成员
            eng.Hide();
            var be4 = new FakeBackend
            {
                Text = MemberSample,
                Extra = new List<ModuleScope> { cls },
                Snap = new Snapshot { Line = 5, Text = "    Me.", Module = "Mod1", CaretCol = 8 },
            };
            eng.Backend = be4;
            Ok("member/Me 用当前模块", eng.Trigger(false, be4.Snap));
        }

        // ------------------------------------------------------------------
        private static void TestEditing()
        {
            // ---- 基础换算 ----
            Eq("indent/4 空格", Editing.IndentWidth("    x"), 4);
            Eq("indent/tab 算 4 列", Editing.IndentWidth("\tx"), 4);
            Eq("leading-ws/混合", Editing.LeadingWs(" \t x"), " \t ");

            // ---- 块结构识别 ----
            Eq("block/If Then", Editing.BlockKind("If x Then"), "if");
            Eq("block/End If 不是块头", Editing.BlockKind("End If"), null);
            Eq("block/For", Editing.BlockKind("    For i = 1 To 3"), "for");
            Eq("block/Do While", Editing.BlockKind("Do While x"), "do");
            Eq("block/While", Editing.BlockKind("While x"), "while");
            Eq("block/With", Editing.BlockKind("With rng"), "with");
            Eq("block/Select Case", Editing.BlockKind("Select Case x"), "select");
            Eq("block/Select 少了 Case", Editing.BlockKind("Select x"), null);
            Eq("block/Private Sub", Editing.BlockKind("Private Sub Foo()"), "sub");
            Eq("block/Static Function", Editing.BlockKind("Static Function F() As Long"), "function");
            Eq("block/Property Get", Editing.BlockKind("Property Get X()"), "property");
            Eq("block/Type", Editing.BlockKind("Type Foo"), "type");
            Eq("block/Enum", Editing.BlockKind("Enum Color"), "enum");
            Eq("block/Else 是分支", Editing.BlockKind("Else"), "branch");
            Eq("block/ElseIf 带 Then", Editing.BlockKind("ElseIf y Then"), "branch");
            Eq("block/ElseIf 没 Then", Editing.BlockKind("ElseIf y"), null);
            Eq("block/Case 是分支", Editing.BlockKind("Case 1"), "branch");
            Eq("block/条件编译", Editing.BlockKind("#If VBA7 Then"), "directive");
            Eq("block/单选 If", Editing.BlockKind("If x Then y = 1"), null);
            Eq("block/Next 不是块头", Editing.BlockKind("Next i"), null);
            Eq("block/普通赋值", Editing.BlockKind("x = 1"), null);
            Eq("block/续行符结尾", Editing.BlockKind("If x Then _"), null);
            Eq("block/字符串里的 Then 不算", Editing.BlockKind("If x = \"a Then\" Then"), "if");
            Eq("block/Rem 整行注释", Editing.BlockKind("Rem If x Then"), null);
            Eq("block/整行注释不是块头", Editing.BlockKind("' For i = 1"), null);

            Eq("closer/For -> Next", Editing.BlockCloser("For i = 1 To 3"), "Next");
            Eq("closer/Do -> Loop", Editing.BlockCloser("Do"), "Loop");
            Eq("closer/With -> End With", Editing.BlockCloser("With rng"), "End With");
            Eq("closer/Else 没有收尾", Editing.BlockCloser("Else"), null);
            Eq("closer/#If 没有收尾", Editing.BlockCloser("#If VBA7 Then"), null);

            Ok("closer-line/Next i", Editing.IsBlockCloserLine("Next i"));
            Ok("closer-line/End If 带注释", Editing.IsBlockCloserLine("End If  '结束"));
            Ok("closer-line/For 不是收尾", !Editing.IsBlockCloserLine("For i = 1 To 3"));

            // ---- 缩进推算 ----
            Eq("indent/块头深一级", Editing.NextLineIndent("    For i = 1 To 3", new List<string>(), "    "),
                "        ");
            Eq("indent/普通行抄自己", Editing.NextLineIndent("    x = 1", new List<string>(), "    "), "    ");
            Eq("indent/注释跟上方第一个非注释行",
                Editing.NextLineIndent("    ' 注释", new List<string> { "Sub A()" }, "    "), "    ");
            Eq("indent/找不到参考行就用自己",
                Editing.NextLineIndent("    ' 注释", new List<string>(), "    "), "    ");
            Eq("indent/Tab 工程跟着用 Tab",
                Editing.NextLineIndent("\tIf x Then", new List<string>(), "\t"), "\t\t");

            // ---- 分支对齐 ----
            string mine, next;
            Ok("branch/Else 与 If 齐平",
                Editing.BranchAlign("    Else", new List<string> { "    If x Then" }, "    ",
                    out mine, out next) && mine == "    " && next == "        ");
            Ok("branch/Case 比 Select Case 深一级",
                Editing.BranchAlign("    Case 1", new List<string> { "    Select Case x" }, "    ",
                    out mine, out next) && mine == "        " && next == "            ");
            Ok("branch/找不到归属者就不动",
                !Editing.BranchAlign("    Else", new List<string>(), "    ", out mine, out next));
            Ok("branch/#Else 与 #If 齐平",
                Editing.BranchAlign("#Else", new List<string> { "#If VBA7 Then" }, "    ",
                    out mine, out next) && mine == "" && next == "    ");

            // ---- 回车接管 ----
            string blockLine = "    For i = 1 To 3";
            Ok("enter/块头行尾接管",
                Editing.EnterIndentWanted(blockLine, Rules.DispCol(blockLine, blockLine.Length)));
            Ok("enter/普通行不接管",
                !Editing.EnterIndentWanted("    x = 1", Rules.DispCol("    x = 1", 9)));
            string comment = "    ' 注释";
            Ok("enter/整行注释接管",
                Editing.EnterIndentWanted(comment, Rules.DispCol(comment, comment.Length)));
            Ok("enter/光标不在行尾不接管", !Editing.EnterIndentWanted(blockLine, 10));
            string tabLine = "\tIf x Then";
            Ok("enter/含 Tab 不接管",
                !Editing.EnterIndentWanted(tabLine, Rules.DispCol(tabLine, tabLine.Length)));
            string wideLine = "    For 每个 = 1 To 10";
            Ok("enter/全角行只认行尾",
                Editing.EnterIndentWanted(wideLine, Rules.DispCol(wideLine, wideLine.Length)) &&
                !Editing.EnterIndentWanted(wideLine, Rules.DispCol(wideLine, wideLine.Length - 3)));

            // ---- 退格 ----
            Ok("bs/整行空白接管", Editing.BlankLineDeleteWanted("    ", 5));
            Ok("bs/整行空白中间不接管", !Editing.BlankLineDeleteWanted("    ", 3));
            Ok("bs/真空行不接管", !Editing.BlankLineDeleteWanted("", 1));
            Ok("bs/有内容的行不接管", !Editing.BlankLineDeleteWanted("    x", 6));

            int left, right;
            Ok("bs-pair/空括号中间接管",
                Editing.PairDeleteSpan("x = f()", Rules.DispCol("x = f()", 6), out left, out right) &&
                left == 5 && right == 6);
            Ok("bs-pair/不在括号中间不接管",
                !Editing.PairDeleteSpan("x = f()", Rules.DispCol("x = f()", 8), out left, out right));
            Ok("bs-pair/括号里有内容不接管",
                !Editing.PairDeleteSpan("f(a)", Rules.DispCol("f(a)", 2), out left, out right));
            Ok("bs-pair/空引号中间接管",
                Editing.PairDeleteSpan("s = \"\"", Rules.DispCol("s = \"\"", 5), out left, out right) &&
                left == 4 && right == 5);

            // ---- 自动配对 ----
            int caretOff;
            bool paired;
            string got = Editing.PairInsertion("abc", 3, '(', ')', out caretOff, out paired);
            Ok("pair/行尾补一对", got == "abc()" && caretOff == 4 && paired);
            got = Editing.PairInsertion("abc def", 0, '(', ')', out caretOff, out paired);
            Ok("pair/右侧还有内容只插左括号", got == "(abc def" && !paired);
            got = Editing.PairInsertion("foo(bar", 7, '(', ')', out caretOff, out paired);
            Ok("pair/实参表里仍补一对", got == "foo(bar()" && paired);
            got = Editing.PairInsertion("s = ", 4, '"', '"', out caretOff, out paired);
            Ok("pair/引号补一对", got == "s = \"\"" && paired);
            got = Editing.PairInsertion("s = \"ab", 7, '"', '"', out caretOff, out paired);
            Ok("pair/未闭合字符串里只补收尾", got == "s = \"ab\"" && !paired);
            got = Editing.PairInsertion("s = \"\"", 5, '"', '"', out caretOff, out paired);
            Ok("pair/右边已是引号就跨过去", got == null && caretOff == 6);
            // 注：`)` 的"跨过去"在 VbeBackend.InsertPair 里先挡掉了（那个分支要读
            //     COM 才知道右边是什么），PairInsertion 不负责它 —— 这里只验引号那一档。

            Eq("pair/全角括号收符号", Editing.CloseOf('（'), '）');
            Eq("pair/全角引号收符号", Editing.CloseOf('“'), '”');
            Eq("pair/不认识的符号", Editing.CloseOf('['), '\0');

            // ---- 配对的闸门：只挡"注释里"，不挡"字符串里" ----
            // 这一条是 `s = "abc|"` 里再按 `"` 别变成 `"abc""` 的前提：
            // 闸门要是把"在字符串里"也算进去，PairInsertion 的"跨过去 / 只补
            // 收尾"两档就永远走不到。
            Ok("gate/未闭合字符串不算注释", !Rules.InCommentOnly("s = \"abc", 9));
            Ok("gate/闭合字符串之后不算注释", !Rules.InCommentOnly("s = \"abc\"", 10));
            Ok("gate/同一行与 InCommentOrString 的差别",
                Rules.InCommentOrString("s = \"abc", 9) && !Rules.InCommentOnly("s = \"abc", 9));
            Ok("gate/引号后是注释", Rules.InCommentOnly("s = \"abc\" ' hi", 14));
            Ok("gate/行首撇号", Rules.InCommentOnly("' hi", 5));
            Ok("gate/字符串里的撇号不是注释", !Rules.InCommentOnly("MsgBox \"it's\"", 13));
            Ok("gate/Rem 整行注释", Rules.InCommentOnly("Rem 说明", 6));
            Ok("gate/Remember 不是注释", !Rules.InCommentOnly("Remember = 1", 11));
            Ok("gate/冒号后的 Rem", Rules.InCommentOnly("x = 1 : Rem 说明", 14));
            Ok("gate/语句中间的 Rem 不是注释", !Rules.InCommentOnly("x = Rem", 7));

            Eq("call-paren/foo( 认得出", Editing.EnclosingCallParen("foo(", 4), 3);
            Eq("call-paren/运算分组 -> -1", Editing.EnclosingCallParen("x = (a + ", 9), -1);
            Eq("call-paren/字符串里的括号不算", Editing.EnclosingCallParen("s = \"(\" & foo(", 14), 13);

            Ok("string/未闭合", Editing.UnterminatedString("x = \"abc"));
            Ok("string/已闭合", !Editing.UnterminatedString("x = \"abc\""));
            Ok("string/注释里的引号不算", !Editing.UnterminatedString("x = 1  ' 他说\"你好"));
            Ok("string/双写引号是转义", !Editing.UnterminatedString("x = \"a\"\"b\""));

            // ---- 补收尾（保守口径）----
            Ok("closer-needed/下面是 End Sub 就补",
                Editing.CloserNeeded(new List<string> { "End Sub" }, "Next", 4));
            Ok("closer-needed/下面已有 Next 不补",
                !Editing.CloserNeeded(new List<string> { "    Next" }, "Next", 4));
            Ok("closer-needed/块里已有内容不补",
                !Editing.CloserNeeded(new List<string> { "        x = 1" }, "Next", 4));
            Ok("closer-needed/同缩进的别的东西不补",
                !Editing.CloserNeeded(new List<string> { "    y = 2" }, "Next", 4));
            Ok("closer-needed/空行与注释跳过",
                Editing.CloserNeeded(new List<string> { "", "    ' 注释", "End Sub" }, "Next", 4));
        }

        // ------------------------------------------------------------------
        // 一键格式化缩进（Ctrl+Shift+L）：纯函数，逐行钉住"应有的行首空白"
        // ------------------------------------------------------------------
        private static void TestIndent()
        {
            // ---- 基础：Sub / If / Else ----
            Eq("fmt/Sub+If+Else", Fmt(
                "Sub A()\nx = 1\nIf x Then\ny = 2\nElse\ny = 3\nEnd If\nEnd Sub"),
                "Sub A()\n    x = 1\n    If x Then\n        y = 2\n    Else\n        y = 3\n    End If\nEnd Sub");

            // ---- Select Case：Case 比 Select 深一级，块体再深一级 ----
            Eq("fmt/Select Case", Fmt(
                "Sub A()\nSelect Case x\nCase 1\ny = 1\nCase Else\ny = 2\nEnd Select\nEnd Sub"),
                "Sub A()\n    Select Case x\n        Case 1\n            y = 1\n        Case Else\n            y = 2\n    End Select\nEnd Sub");

            // ---- 多行块头（续行）：整条语句算一次，续行再深一级 ----
            Eq("fmt/多行 If 头", Fmt(
                "Sub A()\nIf a And _\nb Then\ny = 1\nEnd If\nEnd Sub"),
                "Sub A()\n    If a And _\n        b Then\n        y = 1\n    End If\nEnd Sub");

            // ---- with / for / next ----
            Eq("fmt/With+For", Fmt(
                "Sub A()\nWith rng\nFor i = 1 To 3\nx = i\nNext i\nEnd With\nEnd Sub"),
                "Sub A()\n    With rng\n        For i = 1 To 3\n            x = i\n        Next i\n    End With\nEnd Sub");

            // ---- Do / Loop ----
            Eq("fmt/Do Loop", Fmt("Sub A()\nDo While x\nx = x - 1\nLoop\nEnd Sub"),
                "Sub A()\n    Do While x\n        x = x - 1\n    Loop\nEnd Sub");

            // ---- ElseIf ----
            Eq("fmt/ElseIf", Fmt("Sub A()\nIf a Then\nx = 1\nElseIf b Then\nx = 2\nEnd If\nEnd Sub"),
                "Sub A()\n    If a Then\n        x = 1\n    ElseIf b Then\n        x = 2\n    End If\nEnd Sub");

            // ---- 普通语句的续行 ----
            Eq("fmt/赋值续行", Fmt("Sub A()\nx = a + _\nb\nEnd Sub"),
                "Sub A()\n    x = a + _\n        b\nEnd Sub");

            // ---- 条件编译 ----
            Eq("fmt/#If", Fmt("#If VBA7 Then\nPrivate Declare PtrSafe Sub s Lib \"k\" ()\n#Else\nPrivate Declare Sub s Lib \"k\" ()\n#End If"),
                "#If VBA7 Then\n    Private Declare PtrSafe Sub s Lib \"k\" ()\n#Else\n    Private Declare Sub s Lib \"k\" ()\n#End If");

            // ---- 空行 / 整行注释一个字符都不动 ----
            Eq("fmt/注释与空行原样", Fmt("Sub A()\n  ' 顶格注释\n\nx = 1\nEnd Sub"),
                "Sub A()\n  ' 顶格注释\n\n    x = 1\nEnd Sub");

            // ---- 行标签顶格 ----
            Eq("fmt/标签顶格", Fmt("Sub A()\nIf x Then\nRetry:\nx = 1\nEnd If\nEnd Sub"),
                "Sub A()\n    If x Then\nRetry:\n        x = 1\n    End If\nEnd Sub");

            // ---- Tab 工程：一级就是一个 Tab ----
            Eq("fmt/Tab 单位", Fmt("Sub A()\nx = 1\nEnd Sub", "\t"), "Sub A()\n\tx = 1\nEnd Sub");

            // ---- 已经对齐的：一行都不用改（Changed == 0）----
            var good = Parser.SplitLines("Sub A()\n    x = 1\nEnd Sub");
            Eq("fmt/已对齐不用改", Indent.Compute(good, "    ", 1, good.Length).Changed, 0);

            // ---- 只格式化中间几行：层级仍按全文结构算 ----
            var part = Parser.SplitLines("Sub A()\nIf x Then\ny = 1\nEnd If\nEnd Sub");
            var pr = Indent.Compute(part, "    ", 3, 3);
            Eq("fmt/只改指定行", pr.Changed, 1);
            Eq("fmt/指定行的层级按全文算", pr.Ws[2], "        ");
            Eq("fmt/范围外的行不碰", pr.Ws[1], null);

            // ---- 结构没闭合也不崩：照常算，只是把"末层级"报出来 ----
            var open = Parser.SplitLines("Sub A()\nIf x Then\ny = 1");
            Eq("fmt/未闭合不崩", Indent.Compute(open, "    ", 1, open.Length).FinalDepth, 2);

            // ---- 续行判据 ----
            Ok("fmt/续行符", Indent.IsContinuation("x = 1 _"));
            Ok("fmt/下划线标识符不算续行", !Indent.IsContinuation("x = foo_"));
            Ok("fmt/短行不算续行", !Indent.IsContinuation("_"));
        }

        /// <summary>把 src 按 4 空格（或指定单位）格式化一遍，返回结果文本。</summary>
        private static string Fmt(string src, string unit = "    ")
        {
            var lines = Parser.SplitLines(src);
            var res = Indent.Compute(lines, unit, 1, lines.Length);
            var outp = new List<string>(lines.Length);
            for (int i = 0; i < lines.Length; i++)
            {
                if (res.Ws[i] == null) { outp.Add(lines[i]); continue; }
                int n = Editing.LeadingWs(lines[i]).Length;
                outp.Add(res.Ws[i] + lines[i].Substring(n));
            }
            return string.Join("\n", outp);
        }

        // ------------------------------------------------------------------
        // 自动弹出的总判据（v141）：**行号变了就不能弹**
        // 起因：用户报"有时候鼠标单击到某个变量名，它会自动弹提示词框"。
        // 单击另一行上的变量时，快照的"行号 + 行文本"全变了，和"敲了一个字"
        // 长得一模一样；拦不住它的正是"最近 1.2 秒按过键"那条（打完字顺手点一下）。
        // ------------------------------------------------------------------
        private static void TestAutoTrigger()
        {
            // 正常打字：同行、内容变了、刚按过键 -> 弹
            Ok("auto-trigger/同行打字弹",
                Rules.AutoTriggerWanted(true, false, true, false));
            // 鼠标单击另一行的变量（或方向键上下）：行号跳了 -> 不弹
            Ok("auto-trigger/行号跳了不弹",
                !Rules.AutoTriggerWanted(true, false, true, true));
            // 右键粘贴 / 拖拽 / 别的程序改模块：最近没按键 -> 不弹
            Ok("auto-trigger/没按键不弹",
                !Rules.AutoTriggerWanted(true, false, false, false));
            // 切了代码窗（哪怕刚按过键）-> 不弹
            Ok("auto-trigger/切模块不弹",
                !Rules.AutoTriggerWanted(true, true, true, false));
            // 什么都没变 -> 不弹
            Ok("auto-trigger/没变化不弹",
                !Rules.AutoTriggerWanted(false, false, true, false));
        }

        // ------------------------------------------------------------------
        // 片段库：标题模糊匹配（不碰数据库，纯函数）
        // ------------------------------------------------------------------
        private static void TestSnippets()
        {
            List<int> hits;
            Ok("snippet/整标题相等", SnippetMatch.Score("getcode", "GetCode", out hits) != null);
            Ok("snippet/前缀", SnippetMatch.Score("Get", "GetCode", out hits) != null);
            Ok("snippet/子串", SnippetMatch.Score("Code", "GetCode", out hits) != null);
            Ok("snippet/驼峰子序列 gc -> GetCode", SnippetMatch.Score("gc", "GetCode", out hits) != null);
            Ok("snippet/不命中", SnippetMatch.Score("zzz", "GetCode", out hits) == null);
            Ok("snippet/多词 AND 全中", SnippetMatch.Score("get code", "GetCode", out hits) != null);
            Ok("snippet/多词 AND 有一个不中",
                SnippetMatch.Score("get zzz", "GetCode", out hits) == null);

            int s1 = SnippetMatch.Score("Get", "GetCode", out hits) ?? -1;
            int s2 = SnippetMatch.Score("Code", "GetCode", out hits) ?? -1;
            Ok("snippet/前缀分高于子串分", s1 > s2, "prefix=" + s1 + " substr=" + s2);

            var rows = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("Zeta", "z"),
                new KeyValuePair<string, string>("SetValue", "a"),
                new KeyValuePair<string, string>("Settings", "b"),
            };
            var f = SnippetMatch.Filter(rows, "");
            Eq("snippet/空词原样返回", f.Count, 3);
            Eq("snippet/空词保持原顺序", f[0].Title, "Zeta");

            var f2 = SnippetMatch.Filter(rows, "set");
            Ok("snippet/过滤掉不命中的", f2.Count == 2, "count=" + f2.Count);
            Ok("snippet/前缀命中排前面", f2.Count > 0 && f2[0].Title == "Settings",
                f2.Count > 0 ? f2[0].Title : "");
        }
    }
}
