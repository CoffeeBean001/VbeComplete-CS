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
                TestEditing();
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

            int pos, span;
            Ok("match/前缀", Rules.Match("MsgBox", "msg", out pos, out span) && pos == 0);
            Ok("match/边界跳步", Rules.Match("selBegin", "sb", out pos, out span) && pos == 0 && span == 4);
            Ok("match/非边界不跳步", !Rules.Match("abcDef", "b", out pos, out span) || pos == 0);
            Ok("match/不命中", !Rules.Match("abc", "xyz", out pos, out span));

            Ok("resolved/精确且无更长",
                Rules.PrefixIsResolved("sub", new List<string> { "sub", "function" }));
            Ok("resolved/有更长候选不算打全",
                !Rules.PrefixIsResolved("sub", new List<string> { "sub", "subset" }));
            Ok("resolved/没精确命中",
                !Rules.PrefixIsResolved("su", new List<string> { "sub" }));

            Ok("member-position/点号后", Rules.InMemberPosition("student.", 9));
            Ok("member-position/词中间不是", !Rules.InMemberPosition("student.na", 10));
            Ok("member-position/空格后点号", Rules.InMemberPosition("a = b. ", 8));

            Ok("decl-position/Dim 名字", Rules.InDeclarationPosition("Dim cus", 8));
            Ok("decl-position/As 后是类型位置", !Rules.InDeclarationPosition("Dim x As Inte", 13));
            Ok("decl-position/Sub 名字", Rules.InDeclarationPosition("Sub Do", 7));
            Ok("decl-position/普通语句不是", !Rules.InDeclarationPosition("x = cus", 8));
            Ok("decl-position/Const 值位置", !Rules.InDeclarationPosition("Const PI = 3", 13));

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
                Bank = new WordBank(),
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
