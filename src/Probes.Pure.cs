using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace VbeComplete
{
    /// <summary>离线自检：片段库、词库、完整判定链。</summary>
    internal static partial class Program
    {
        /// <summary>
        /// 片段库自检（不需要 Excel、不需要桌面）：在**临时库**上跑一遍增删改查，
        /// 顺便报告真实库的路径与条数。
        /// ★刻意不碰用户的 snippets-cs.db —— 自检绝不允许改坏用户攒下来的片段。
        /// </summary>
        private static int SnippetTest()
        {
            var sb = new StringBuilder();
            bool ok = true;
            string tmp = Path.Combine(Path.GetTempPath(),
                "vbecomplete-snipetest-" + Guid.NewGuid().ToString("N") + ".db");

            sb.AppendLine("=== snippet store test ===");
            sb.AppendLine("winsqlite3 路径 = " + Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System), "winsqlite3.dll"));
            sb.AppendLine("临时库 = " + tmp);
            sb.AppendLine("真实库 = " + SnippetStore.DbPath() +
                          "（存在=" + File.Exists(SnippetStore.DbPath()) + "）");

            Action<string, bool, string> check = (name, pass, detail) =>
            {
                if (!pass) ok = false;
                sb.AppendLine((pass ? "  PASS  " : "  FAIL  ") + name + (detail.Length > 0 ? "   " + detail : ""));
            };

            // 走环境变量换库（Settings.Raw 每次都现读，所以这里设了就生效）
            string oldEnv = Environment.GetEnvironmentVariable("VBECOMPLETE_SNIPPET_DB");
            Environment.SetEnvironmentVariable("VBECOMPLETE_SNIPPET_DB", tmp);
            try
            {
                using (var st = new SnippetStore())
                {
                    check("打开临时库", st.Count() == 0, "count=" + st.Count() + " err=" + st.LastError());

                    check("新增", st.Add("Hello", "MsgBox \"hi\""), "err=" + st.LastError());
                    check("新增第二条", st.Add("Loop", "For i = 1 To 3\r\nNext"), "");
                    check("重名新增返回 false", !st.Add("Hello", "x"), "");
                    check("读回来一致", st.Get("Hello") == "MsgBox \"hi\"", "got=[" + st.Get("Hello") + "]");
                    check("Exists", st.Exists("Loop") && !st.Exists("nope"), "");
                    check("修改代码", st.Update("Hello", "MsgBox \"bye\"") &&
                          st.Get("Hello") == "MsgBox \"bye\"", "");

                    var all = st.All();
                    check("全部按标题排序", all.Count == 2 && all[0].Key == "Hello" && all[1].Key == "Loop",
                        "count=" + all.Count);

                    check("改名", st.Rename("Hello", "Hi") && st.Get("Hi") != null && !st.Exists("Hello"), "");
                    check("改名撞车返回 false", !st.Rename("Hi", "Loop"), "");

                    check("中文标题与多行代码",
                        st.Add("中文片段", "Sub 测试()\r\n    Debug.Print \"中文\"\r\nEnd Sub") &&
                        (st.Get("中文片段") ?? "").Contains("中文"), "");

                    check("删除", st.Delete("Hi") && !st.Exists("Hi"), "");
                    check("删除不存在返回 false", !st.Delete("Hi"), "");
                }

                // 重新打开一次：确认是真的落盘了（不是内存里的假象）
                using (var st2 = new SnippetStore())
                {
                    check("重新打开后数据还在", st2.Count() == 2, "count=" + st2.Count());
                }
            }
            catch (Exception e)
            {
                ok = false;
                sb.AppendLine("EXCEPTION " + e);
            }
            finally
            {
                Environment.SetEnvironmentVariable("VBECOMPLETE_SNIPPET_DB", oldEnv);
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
            }

            sb.AppendLine(ok ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "snipetest.log"), sb.ToString(), Encoding.UTF8);
            return ok ? 0 : 1;
        }

        /// <summary>
        /// 词库自检（不需要 Excel、不需要桌面）：在**临时库**上验证
        /// "第一次自动建库并灌种子 → 载入内存 → 用户加一行也能生效"。
        /// ★刻意不碰真实的关键字/函数库 —— 自检绝不允许改坏用户自己维护的词库。
        /// </summary>
        private static int DbTest()
        {
            var sb = new StringBuilder();
            bool ok = true;
            string tmpKw = Path.Combine(Path.GetTempPath(),
                "vbecomplete-kw-" + Guid.NewGuid().ToString("N") + ".db");
            string tmpFn = Path.Combine(Path.GetTempPath(),
                "vbecomplete-fn-" + Guid.NewGuid().ToString("N") + ".db");

            sb.AppendLine("=== word db test ===");
            sb.AppendLine("真实关键字库 = " + WordLibrary.KeywordPath() +
                          "（存在=" + File.Exists(WordLibrary.KeywordPath()) + "）");
            sb.AppendLine("真实函数库   = " + WordLibrary.FunctionPath() +
                          "（存在=" + File.Exists(WordLibrary.FunctionPath()) + "）");
            sb.AppendLine("临时关键字库 = " + tmpKw);
            sb.AppendLine("临时函数库   = " + tmpFn);

            Action<string, bool, string> check = (name, pass, detail) =>
            {
                if (!pass) ok = false;
                sb.AppendLine((pass ? "  PASS  " : "  FAIL  ") + name +
                              (string.IsNullOrEmpty(detail) ? "" : "   " + detail));
            };

            string oldKw = Environment.GetEnvironmentVariable("VBECOMPLETE_KEYWORD_DB");
            string oldFn = Environment.GetEnvironmentVariable("VBECOMPLETE_FUNCTION_DB");
            Environment.SetEnvironmentVariable("VBECOMPLETE_KEYWORD_DB", tmpKw);
            Environment.SetEnvironmentVariable("VBECOMPLETE_FUNCTION_DB", tmpFn);
            try
            {
                var lib = new WordLibrary();
                lib.Load();
                sb.AppendLine("载入结果：关键字 " + lib.KeywordCount + " 条，函数/类型 " +
                              lib.FunctionCount + " 条，" + lib.LastError);
                check("第一次运行自动建库", File.Exists(tmpKw) && File.Exists(tmpFn), "");
                check("关键字灌了种子", lib.KeywordCount >= 70, "count=" + lib.KeywordCount);
                check("函数灌了种子", lib.FunctionCount >= 150, "count=" + lib.FunctionCount);
                check("能查到 Dim", VbaBuiltins.IsKeyword("Dim"), "");
                check("能查到 MsgBox", VbaBuiltins.IsBuiltinFunction("MsgBox"), "");
                check("能查到 String（类型）", VbaBuiltins.IsBuiltinType("String"), "");
                check("枚举一律不收（vbRed）", !VbaBuiltins.IsBuiltin("vbRed"), "");
                check("枚举一律不收（xlUp）", !VbaBuiltins.IsBuiltin("xlUp"), "");
                check("类型名不补括号", !VbaBuiltins.IsCallableBuiltin("String"), "");
                check("函数补括号", VbaBuiltins.IsCallableBuiltin("Split"), "");
                check("Me 不补空格", !VbaBuiltins.KeywordWantsSpace("Me"), "");
                check("Dim 补空格", VbaBuiltins.KeywordWantsSpace("Dim"), "");
                check("Next 不补空格", !VbaBuiltins.KeywordWantsSpace("Next"), "");

                // 用户自己往里加一行 —— 重启后必须能进候选池
                bool inserted;
                using (var db = new SqliteDb())
                {
                    inserted = db.Open(tmpFn) &&
                               db.Write("INSERT OR REPLACE INTO functions " +
                                        "(word, kind, wants_parens, enabled) " +
                                        "VALUES ('我的函数', 'function', 1, 1)");
                }
                check("往库里插一行", inserted, "");
                new WordLibrary().Load();
                check("自定义词条生效", VbaBuiltins.IsBuiltinFunction("我的函数"), "");
                check("自定义词条补括号", VbaBuiltins.IsCallableBuiltin("我的函数"), "");

                // enabled=0 的条目必须消失
                using (var db = new SqliteDb())
                {
                    if (db.Open(tmpFn))
                        db.Write("UPDATE functions SET enabled=0 WHERE word='我的函数'");
                }
                new WordLibrary().Load();
                check("enabled=0 的词条不提示", !VbaBuiltins.IsBuiltinFunction("我的函数"), "");

                // ---- v175：设置界面用的那层（WordBank）—— 增 / 改 / 停用 / 删 ----
                // ★这一段就是在验"用户点『新增』之后到底发生了什么"：写库 + 重灌内存。
                //   全程跑在临时库上（上面已经把两个环境变量指到临时文件了）。
                string err;
                check("WordBank 加关键字",
                    WordBank.Save(WordBank.Keywords,
                        new WordEntry { Word = "我的关键字", Enabled = true, WantsSpace = true },
                        null, out err), err);
                var kw1 = WordBank.All(WordBank.Keywords, out err);
                check("WordBank 读回关键字", kw1.Any(x => x.Word == "我的关键字"), err);

                WordBank.Reload();
                check("加完立刻进候选池（不用重启）", VbaBuiltins.IsKeyword("我的关键字"), "");

                check("WordBank 改名（主键是 word，删+插）",
                    WordBank.Save(WordBank.Keywords,
                        new WordEntry { Word = "我的关键字2", Enabled = true, WantsSpace = false },
                        "我的关键字", out err), err);
                kw1 = WordBank.All(WordBank.Keywords, out err);
                check("改名后旧词没了", !kw1.Any(x => x.Word == "我的关键字"), "");
                check("改名后新词在", kw1.Any(x => x.Word == "我的关键字2"), "");

                check("WordBank 停用",
                    WordBank.SetEnabled(WordBank.Keywords, new[] { "我的关键字2" }, false, out err), err);
                WordBank.Reload();
                check("停用后不进候选", !VbaBuiltins.IsKeyword("我的关键字2"), "");

                int del = WordBank.Delete(WordBank.Keywords, new[] { "我的关键字2" }, out err);
                check("WordBank 删除（返回真的删掉几条）", del == 1, "del=" + del + " err=" + err);
                kw1 = WordBank.All(WordBank.Keywords, out err);
                check("删掉的词不在库里", !kw1.Any(x => x.Word == "我的关键字2"), "");

                check("WordBank 加类型（kind=type）",
                    WordBank.Save(WordBank.Functions,
                        new WordEntry { Word = "我的类型", Kind = "type", Enabled = true,
                            WantsParens = true }, null, out err), err);
                WordBank.Reload();
                check("类型名不补括号", VbaBuiltins.IsBuiltinType("我的类型") &&
                                        !VbaBuiltins.IsCallableBuiltin("我的类型"), "");

                string why;
                check("带空格的词条被拒", !WordBank.IsValidWord("a b", out why), why);
                check("纯数字的词条被拒", !WordBank.IsValidWord("123", out why), why);
                check("中文标识符合法", WordBank.IsValidWord("我的函数", out why), why);
            }
            catch (Exception e)
            {
                ok = false;
                sb.AppendLine("EXCEPTION " + e);
            }
            finally
            {
                Environment.SetEnvironmentVariable("VBECOMPLETE_KEYWORD_DB", oldKw);
                Environment.SetEnvironmentVariable("VBECOMPLETE_FUNCTION_DB", oldFn);
                VbaBuiltins.ResetToSeed();
                try { if (File.Exists(tmpKw)) File.Delete(tmpKw); } catch { }
                try { if (File.Exists(tmpFn)) File.Delete(tmpFn); } catch { }
            }

            sb.AppendLine(ok ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "dbtest.log"), sb.ToString(), Encoding.UTF8);
            return ok ? 0 : 1;
        }

        /// <summary>
        /// 候选判定的完整回归：真实模块文本 + 人造快照跑 Trigger，逐条验"弹不弹、第一条是什么"。
        /// ★Trigger 的全部输入就是 (模块文本, 行, 列)，人造快照走同一条代码路径，判定逻辑
        ///   100% 覆盖；"按键有没有送到钩子"由 --hktest 兜。纯读，一个字符都不写模块。
        /// </summary>
        private static int EngineTest()
        {
            var sb = new StringBuilder();
            bool ok = true;

            // 真机信息只打印、不作判据 —— 没开代码窗时这里是 -1。
            // ★旧实现把整节押在"真机模块里恰好有 student2 / okkk2 / 他们"上，
            //   换个模块就整片红或整片 skip；那样的绿说明不了任何事
            //   （v137 之前 paretest / hktest / enginetest 三条全是这么"绿"着的）。
            string liveText = null;
            int liveLines = -1;
            try
            {
                var live = new VbeBackend();
                liveText = live.ModuleText();
                liveLines = live.LineCount();
            }
            catch { }

            sb.AppendLine("=== engine test（合成工程，不依赖真机打开的模块）===");
            sb.AppendLine("真机模块（仅供参考）：lines=" + liveLines +
                          " len=" + (liveText == null ? -1 : liveText.Length));

            // 合成工程：一个标准模块（模块1）+ 一个类模块（cls1）。
            // 组件类型要真的填（1=标准 2=类）—— "类成员不进裸名池"就靠它判。
            const string modStd =
                "Option Explicit\n" +
                "\n" +
                "Public 我们 As String\n" +
                "Private okkk2 As String\n" +
                "\n" +
                "Sub test()\n" +
                "    Dim student2 As String\n" +
                // ★v173：`resVar` 只用来看"纯语法位置会不会把名字也漏出来"（用户截图里
                //   那条 `result` 是同一类）。**必须并进这一行、不能新起一行** —— 下面
                //   "去掉 Option Explicit 的 A/B 对照"写死了 Line = 11，加行会把它顶偏，
                //   那条就变成假红（v173 踩过一次）。
                "    Dim 他们 As String, resVar As String\n" +
                "    Dim dic As Object, c As cls1\n" +
                "    student = getOne()\n" +
                "    \n" +
                "End Sub\n";
            const string modCls =
                "Option Explicit\n" +
                "\n" +
                "Public s1 As String\n" +
                "Public s2 As String\n";
            // ★v146：故意让这个模块里只有 `wsR.Cells(...)` 这种**成员访问**。
            //   旧版会把它拆成"wsR"和"Cells"两个自由标识符，Cells 就此进了全工程
            //   的裸名池 —— 用户在**任何**模块打 `cell` 都会冒出 `Cells`。
            const string modMem =
                "Sub W()\n" +
                "    n = wsR.Cells(1, 1).Value\n" +
                "End Sub\n";
            // ★v153：无 Option Explicit 的 Worksheet 模块（用户报的那一段原文）。
            //   `Range` 只在两处出现：`As Range`（类型）与 `Range("d2")`（调用），
            //   都不是"自动声明的变量" —— 打 `ra` 不该冒出 Range。
            const string modNoExp =
                "Private Sub Worksheet_SelectionChange(ByVal Target As Range)\n" +
                "    Set Target = Target.Cells(1, 1)\n" +
                "    If Target.Address = Range(\"d2\").Address Then\n" +
                "        setValidity ActiveCell, \"c\"\n" +
                "    End If\n" +
                "    ra\n" +
                "    n7 = Selection.End(xlUp)\n" +
                "End Sub\n";

            // ★v154：无 Option Explicit 的隐式变量是**模块私有**的 —— 另一个标准模块
            //   里写过的 `secretVar` 不该在别处冒出来。模块2 里同时放一行 `sec`，
            //   用来证明"在自己的模块里照常提示"。
            const string modFree = "Sub Q()\n" +
                                   "    secretVar = 1\n" +
                                   "    sec\n" +
                                   "End Sub\n";
            // Type 块里的成员（只能 p.Nickname 访问）不是裸名
            const string modType = "Public Type Person\n" +
                                   "    Nickname As String\n" +
                                   "    Age As Long\n" +
                                   "End Type\n";
            // ★v158：用户报的原文（无 Option Explicit 的 Worksheet 模块）：
            //   `Set conditions = getConditions()` 之后，在 `If isMeet(condi|` 里
            //   打 condi 提示不出 conditions。
            //   ★关键：下面 `Private Function isMeet(conditions, ws, t, r)` 的**形参**
            //   也叫 conditions —— 旧实现按名字一刀切排除，形参把隐式变量挡掉了。
            const string modUser =
                "Sub 查询()\n" +
                "    Set ws = Worksheets(\"人员考勤记录\")\n" +
                "    Set t = getTitle(ws, 1, 1)\n" +
                "    Set conditions = getConditions()      ' 已设置的条件\n" +
                "    Set dic = CreateObject(\"scripting.dictionary\")\n" +
                "    For i = 2 To ws.Cells(Cells.Rows.Count, t(\"姓名\")).End(xlUp).Row\n" +
                "        If isMeet(condi\n" +
                "    Next\n" +
                "End Sub\n" +
                "\n" +
                "Private Function getConditions()\n" +
                "    Set dic = CreateObject(\"scripting.dictionary\")\n" +
                "    Set getConditions = dic\n" +
                "End Function\n" +
                "\n" +
                "Private Function isMeet(conditions, ws, t, r)\n" +
                "    If conditions.exists(\"项目名称\") Then\n" +
                "    End If\n" +
                "End Function\n";

            // ★v160：用户报的第二个案例（无 Option Explicit）：
            //   `Set dic = CreateObject(...)` / `dic.Add s, i` 之后，在
            //   `set getTitle2=di|` 里打 di 提示不出 dic。
            const string modUser2 =
                // ★v161：用户真实模块里 getTitle 和 getTitle2 **都有** `Set dic = ...`
                //   —— 同名隐式变量在两个过程里各有一份，只记第一次出现的那个过程，
                //   getTitle2 里就再也看不见它了。
                "Function getTitle(ws, r1)\n" +
                "    Set dic = CreateObject(\"scripting.dictionary\")\n" +
                "    Set getTitle = dic\n" +
                "End Function\n" +
                "\n" +
                "Function getTitle2(ws, r1, r2)\n" +
                "    Set dic = CreateObject(\"scripting.dictionary\")\n" +
                "    lastCol1 = ws.Cells(r1, Cells.Columns.Count).End(xlToLeft).Column\n" +
                "    lastCol2 = ws.Cells(r2, Cells.Columns.Count).End(xlToLeft).Column\n" +
                "    lastCol = Application.WorksheetFunction.Max(lastCol1, lastCol2)\n" +
                "    For i = 1 To lastCol\n" +
                "        s1 = ws.Cells(r1, i).MergeArea.Cells(1, 1)\n" +
                "        s2 = ws.Cells(r2, i).MergeArea.Cells(1, 1)\n" +
                "        s = s1 & s2\n" +
                "        dic.Add s, i\n" +
                "    Next\n" +
                "    set getTitle2=di\n" +
                "End Function\n";

            // ★v176：用户报"写了 for i，这里的 i 为什么不提示" —— 代码里是
            //   `Dim i&, j&, lastCol&`（带类型后缀、无 As）。单独一个模块，因为
            //   要的行号很短，混进模块1 会被上面那些用例的行号顶偏。
            const string modFor =
                "Option Explicit\n" +
                "\n" +
                "Sub 循环()\n" +
                "    Dim i&, j&, lastCol&\n" +
                "    For i\n" +
                "End Sub\n";

            // ★v178：形参表跨续行符（用户截图里的原文）。两份：不带 / 带 Option Explicit。
            //   带 Option Explicit 的那份正是 bug 现场 —— 形参一条都收不到，"用过但没声明"
            //   那条路又被 Option Explicit 堵死，于是打 di 只剩内建（Dir / Dim / ReDim …）。
            //   第 10 行 `if di` 与第 11 行 `co` 是两条探针的落点（加了 Option Explicit 各 +1）。
            const string modParamNoExp =
                "'获取公式\n" +
                "Private Function getFormula(ByVal ws As Worksheet, ByVal col As Long, _\n" +
                "                            ByVal dic1 As Object, ByVal dic2 As Object) As String\n" +
                "    Dim i&, s1$, s2$, rate\n" +
                "    For i = 2 To ws.Cells(ws.Rows.Count, \"b\").End(xlUp).Row\n" +
                "        s1 = ws.Cells(i, \"a\").MergeArea.Cells(1, 1)\n" +
                "        s2 = ws.Cells(i, \"b\")\n" +
                "        If s1 = \"门诊\" Then\n" +
                "            if di\n" +
                "            co\n" +
                "        End If\n" +
                "    Next\n" +
                "End Function\n";
            const string modParamExp = "Option Explicit\n" + modParamNoExp;

            // ★v178：`Dim` 也跨续行 —— 续行上的名字同样是声明，不能丢。
            //   第 2-3 行是模块级、第 6-7 行是过程内，两种作用域各验一条。
            const string modDim =
                "Option Explicit\n" +
                "Dim gOne As Long, _\n" +
                "    gTwo As String\n" +
                "\n" +
                "Sub 取数()\n" +
                "    Dim aaa As Long, _\n" +
                "        bbb As String\n" +
                "    aaa = 1\n" +
                "    bb\n" +
                "    gT\n" +
                "End Sub\n";

            var synth = new SyntheticBackend()
                .Add("模块1", 1, modStd)
                .Add("cls1", 2, modCls)
                .Add("modMem", 1, modMem)
                .Add("Sheet1", 100, modNoExp)
                .Add("模块2", 1, modFree)
                .Add("modT", 1, modType)
                .Add("模块3", 1, modUser)
                .Add("模块4", 1, modUser2)
                .Add("模块5", 1, modFor)
                .Add("modParam", 1, modParamNoExp)
                .Add("modParamX", 1, modParamExp)
                .Add("modDim", 1, modDim)
                // v183：窗体。控件名只存在于设计器数据里，真机由 Designer 灌进来，
                //   这里直接摆（验的是"灌进来之后该在哪个模块可见"）。
                .AddForm("UserForm1",
                    "Private Sub CommandButton1_Click()\n" +
                    "    Dim s$\n" +
                    "    s = te\n" +
                    "End Sub\n",
                    "TextBox1", "CommandButton1", "Label1")
                // 一行代码都没有的新窗体：控件名照样要能提示（读代码时不能因为
                // 行数为 0 就把整个组件跳过）
                .AddForm("UserForm2", "", "ListBox1", "SpinButton1");

            // 说明, 模块, 行, 行文本, 光标显示列, 期望该弹?, 必须有, 不许有
            Action<string, string, int, string, int, bool?, string[], string[]> sprobe =
                (name, mod, line, lineText, col, wantShow, mustHave, mustNot) =>
                {
                    synth.Current = mod;
                    var e = new Engine { Recent = new Recent(), Backend = synth };
                    var sn = new Snapshot
                    {
                        Line = line,
                        Text = lineText,
                        Module = mod,
                        CaretCol = col
                    };
                    e.InvalidatePool();
                    bool shown = e.Trigger(false, sn);
                    var items = shown ? e.Items.ToList() : new List<Candidate>();

                    bool pass = true;
                    if (wantShow.HasValue && shown != wantShow.Value) pass = false;
                    foreach (var want in mustHave ?? Array.Empty<string>())
                        if (!items.Any(x => string.Equals(x.Name, want, StringComparison.OrdinalIgnoreCase)))
                            pass = false;
                    foreach (var bad in mustNot ?? Array.Empty<string>())
                        if (items.Any(x => string.Equals(x.Name, bad, StringComparison.OrdinalIgnoreCase)))
                            pass = false;
                    if (!pass) ok = false;

                    sb.AppendLine((pass ? "  PASS  " : "  FAIL  ") + name);
                    sb.AppendLine("         行[" + (lineText ?? "") + "] 列=" + col + " -> " +
                                  (shown ? string.Join(",", items.Take(10).Select(x => x.Name)) : "(不弹)") +
                                  (mustHave != null && mustHave.Length > 0
                                      ? "   必须有=" + string.Join("/", mustHave) : "") +
                                  (mustNot != null && mustNot.Length > 0
                                      ? "   不许有=" + string.Join("/", mustNot) : ""));
                    if (shown) e.Hide();
                };

            // ---- 1. 声明位置 / `As` 之后不弹 ----
            sprobe("As 后不弹（Dim student2 As Str|）", "模块1", 7,
                "    Dim student2 As String", 24, false, null, null);
            sprobe("裸声明名字位不弹（Private okkk2|）", "模块1", 4,
                "Private okkk2 As String", 14, false, null, null);
            sprobe("裸声明 + As 不弹（Public 我们 As Str|）", "模块1", 3,
                "Public 我们 As String", 19, false, null, null);
            sprobe("裸声明名字位不弹（Public 我|）", "模块1", 3,
                "Public 我们 As String", 10, false, null, null);

            // ---- 1b. v168：声明位置里的**第二个关键字**照弹 ----
            //   用户对着 `Private Function` 和 `ReDim Preserve` 各框了一处："这两处
            //   一个提示都没有"。声明位置不提示变量名是对的，但不该连关键字一起封掉。
            sprobe("Private 后补关键字（Private Fu|）", "模块1", 4,
                "Private Fu", 11, true, new[] { "Function" },
                new[] { "Public", "Property", "Preserve" });
            sprobe("ReDim 后补 Preserve（ReDim Pre|）", "模块1", 4,
                "    ReDim Pre", 14, true, new[] { "Preserve" },
                new[] { "Private", "Public" });
            // 同一个位置的名字照旧不弹（`Private m_arr` 那条已在上面）
            sprobe("Private 后的变量名不弹（Private ok|）", "模块1", 4,
                "Private ok", 11, false, null, null);

            // ---- 1c. Then/Else 也是语句边界 + "关键字后跟关键字" ----
            //   ★只认"行的第一个词"时，`If n > 0 Then ReDim Preserve res` 里的 ReDim 不在
            //     行首，"声明位置不弹"整条失效。
            //   ★`ReDim` 命名的是**既有**数组，这里该照弹名字。
            sprobe("单行 If 里 Preserve 后照弹名字（okk|）", "模块1", 4,
                "    If n > 0 Then ReDim Preserve okk", 37, true, new[] { "okkk2" }, null);
            sprobe("ReDim Preserve 后照弹名字（okk|）", "模块1", 4,
                "    ReDim Preserve okk", 23, true, new[] { "okkk2" },
                // ★v170：`Reset` 不在关键字库里（在语句表里），光判 IsKeyword 会漏掉它
                new[] { "Reset", "Resume" });
            sprobe("ReDim 后照弹名字（okk|）", "模块1", 4,
                "    ReDim okk", 14, true, new[] { "okkk2" }, null);
            sprobe("Dim 括号里照弹（0 To okk|）", "模块1", 4,
                "    Dim arr2(0 To okk", 22, true, new[] { "okkk2" }, null);
            // 但 `Dim` 自己的名字位照旧不弹（别修过头）
            sprobe("Dim 名字位仍不弹（Dim okk|）", "模块1", 4,
                "    Dim okk", 12, false, null, null);
            sprobe("单行 If 里 Then 后仍给 Preserve", "模块1", 5,
                "    If n > 0 Then ReDim Pre", 28, true, new[] { "Preserve" }, null);
            sprobe("End 后给 End 的子关键字（End Su|）", "模块1", 5,
                "    End Su", 11, true, new[] { "Sub" }, null);
            sprobe("On Error 后给 GoTo（On Error Go|）", "模块1", 5,
                "    On Error Go", 15, true, new[] { "GoTo" }, null);
            sprobe("For 后给 Each（For Ea|）", "模块1", 5,
                "    For Ea", 11, true, new[] { "Each" }, null);
            sprobe("Case 后给 Else（Case Els|）", "模块1", 5,
                "    Case Els", 13, true, new[] { "Else" }, null);
            // 关键字过滤：`Option Ex|` 只该有 Explicit，不该把同前缀的 Exit 也算进来
            sprobe("Option 后剔掉位置不对的关键字（Option Ex|）", "模块1", 5,
                "    Option Ex", 14, true, new[] { "Explicit" }, new[] { "Exit" });

            // ---- 1d. v171：位置判据不许外溢到"同一句话里的普通表达式" ----
            //   ★用户报：行首打 `Preserve re` 只剩一个 `result`，`ReDim` 自己被滤掉了
            //     —— `Preserve` 只在 `ReDim` 后面才是"只能跟数组名"的位置。
            //   ★v174：这两条的行号从 5 挪到 11 —— 第 5 行是**模块级**，而 `ReDim` 只能
            //     写在过程体里（Rules.KeywordAllowed 会把它剔掉，见下面 1d 那组）。
            sprobe("行首 Preserve 后仍给 ReDim（Preserve re|）", "模块1", 11,
                "    Preserve re", 16, true, new[] { "ReDim" }, null);
            //   ★v173：用户接着报"应该只提示 redim 才对吧" —— 对，位置对了一半不算对。
            //     这个位置是纯语法位，除了 ReDim 什么都不合法：名字（resVar）和同前缀的
            //     内建/关键字（Replace/Rem/Reset/Resume/Return/RaiseEvent）全得剔。
            sprobe("行首 Preserve 只给 ReDim（名字内建全剔）", "模块1", 11,
                "    Preserve re", 16, true, new[] { "ReDim" },
                new[] { "resVar", "Replace", "Rem", "Reset", "Resume", "Return", "RaiseEvent" });
            //   `End Su|` 同理：`End` 后面只可能是关键字，名字一律不合法
            sprobe("End 后不给名字（End Su|）", "模块1", 5,
                "    End Su", 11, true, new[] { "Sub" }, new[] { "resVar" });
            //   ★`Case` 后面是表达式，`Chr` 这类内建不许被当成"位置不对的关键字"剔掉
            sprobe("Case 后给内建函数（Case Ch|）", "模块1", 5,
                "    Case Ch", 12, true, new[] { "Chr" }, null);
            //   ★`Declare PtrSafe Fu|` 是第三个关键字位置（和 `Private Fu|` 同形）
            sprobe("PtrSafe 后给 Function（PtrSafe Fu|）", "模块1", 5,
                "    Private Declare PtrSafe Fu", 31, true, new[] { "Function" }, null);
            //   ★v172：这一格是**真机 declprobe 报的红**（`x = 1: ReDim Preserve okk|`
            //     全哑），但查下来判据本身是对的 —— 离线跑、单跑这一条都绿，红的是
            //     探针固定 sleep 400ms 采样太早（见用例循环里那段说明）。留着当回归线：
            //     冒号切出来的语句段必须和单行 If 的 `Then` 之后走同一套口径。
            sprobe("冒号后面 Preserve 照弹名字（x = 1: …okk|）", "模块1", 11,
                "    x = 1: ReDim Preserve okk", 30, true, new[] { "okkk2" },
                new[] { "Reset", "Resume" });
            sprobe("冒号后面 ReDim 照弹名字（x = 1: ReDim okk|）", "模块1", 11,
                "    x = 1: ReDim okk", 21, true, new[] { "okkk2" }, null);

            // ---- 1d. v174：关键字的使用场合（用户报：标准模块里冒出 Property、
            //   函数体里冒出 Private）。模块1 是标准模块（Type=1），第 5 行是模块级空行，
            //   第 11 行在 `Sub test()` 里面；cls1 是类模块（Type=2）。
            sprobe("标准模块的模块级不给 Property（Prop|）", "模块1", 5,
                "    Prop", 9, false, null, new[] { "Property" });
            sprobe("类模块的模块级有 Property（cls1 里 Prop|）", "cls1", 3,
                "    Prop", 9, true, new[] { "Property" }, null);
            sprobe("函数体里不给 Private（Pri|）", "模块1", 11,
                "    Pri", 8, null, null, new[] { "Private" });
            sprobe("模块级有 Private（Pri|）", "模块1", 5,
                "    Pri", 8, true, new[] { "Private" }, null);
            sprobe("模块级不给 ReDim（ReDi|）", "模块1", 5,
                "    ReDi", 9, false, null, new[] { "ReDim" });
            sprobe("过程体里有 ReDim（ReDi|）", "模块1", 11,
                "    ReDi", 9, true, new[] { "ReDim" }, null);

            // ---- 1e. v176：单字母变量名（`Dim i&`）在循环头要提示 ----
            //   用户报"写了 for i，这里的 i 为什么不提示，单个字符的变量应该会提示自身"。
            sprobe("For 后给单字母变量（For i|）", "模块5", 5,
                "    For i", 10, true, new[] { "i" }, null);

            // ---- 1f. v178：跨续行符的声明（形参表 / Dim）----
            //   用户报："这个 if 后面也不会提示变量名，不能提示 dic1、dic2，应该跟上面的
            //   for 那个是同样问题"。**根因不同**：for 那条是"打全即收"，这条是候选池里
            //   压根没有这个名字 —— `Private Function f(a, b, _` 的形参表跨了续行符，
            //   `ReProcHead` 只吃一行，第一行的 `(` 配不到 `)`，整条参数表一条都没收下来。
            sprobe("续行形参：if 后打 di 给 dic1/dic2（Option Explicit）", "modParamX", 10,
                "            if di", 18, true, new[] { "dic1", "dic2" }, null);
            sprobe("续行形参：无 Option Explicit 也给", "modParam", 9,
                "            if di", 18, true, new[] { "dic1", "dic2" }, null);
            // 第一行的形参同样丢过（ws 只是碰巧因为 `ws.Cells` 像宿主对象才混进来）
            sprobe("续行形参：第一行的 col 也要有（co|）", "modParamX", 11,
                "            co", 15, true, new[] { "col" }, null);
            sprobe("Dim 跨续行：续行上的局部队名要给（bb| -> bbb）", "modDim", 9,
                "    bb", 7, true, new[] { "bbb" }, null);
            sprobe("Dim 跨续行：续行上的模块级名要给（gT| -> gTwo）", "modDim", 10,
                "    gT", 7, true, new[] { "gTwo" }, null);

            // ---- 1g. v180：形参表里"一个新参数的开头"要提示 ByVal / ByRef ----
            //   用户报："在形参列表里，不会提示 byval 和 byref"。整行是 Sub 声明 ⇒ 被判成
            //   "正在起一个新名字"（参数名确实是新名字）⇒ 一个候选都不出。这一格只该收窄成
            //   四个修饰词 + 参数名：`Boolean` 这类类型名不能出现在参数名之前。
            sprobe("形参开头给 ByVal/ByRef（Sub Foo(…, B|）", "模块1", 11,
                "    Sub Foo(a As Long, B", 25, true, new[] { "ByVal", "ByRef" },
                new[] { "Boolean", "Byte" });
            // 参数名本身仍然不弹（v169 的规矩）：`…, okk|` 不是那四个词的前缀
            sprobe("形参里打参数名仍不弹（…, okk|）", "模块1", 11,
                "    Sub Foo(a As Long, okk", 27, false, null, null);
            // 反例：同一行 `As` 之后仍然是类型位置（不弹），别被上面那条放宽过头
            sprobe("形参 As 后照旧不弹（…, b As Str|）", "模块1", 11,
                "    Sub Foo(a As Long, b As Str", 32, false, null, null);

            // ---- 2. Option Explicit：未声明的名字不进候选 ----
            // 光标所在行会被"挖掉半截词"，所以搜索词要和模块里的名字错开：
            // 池子里若混进了第 10 行的 `student`，它就会在 stud 这次冒出来。
            sprobe("未声明 student 不进候选（stud|）", "模块1", 11,
                "    stud", 9, null, new[] { "student2" }, new[] { "student" });
            sprobe("未声明 ok 不进候选（ok|）", "模块1", 11,
                "    ok", 7, null, new[] { "okkk2" }, new[] { "ok" });
            sprobe("中文变量名能提示（他|）", "模块1", 11,
                "    他", 7, true, new[] { "他们" }, null);

            // ---- 2b. v153：无 Option Explicit 时，"只当函数调用过"的名字不算变量 ----
            //   用户报的原文：Worksheet 模块里打 `ra`，第一条是 Excel 的 `Range` ——
            //   它只是代码里 `Range("d2")` 调用过一次的宿主函数。
            sprobe("无 Option Explicit 时调用名不进变量池（ra| 不冒 Range）", "Sheet1", 6,
                "    ra", 7, true, new[] { "RaiseEvent" }, new[] { "Range" });
            // v155：宿主枚举常量（xlUp / vbYes / mso*）不进池 —— PY 的 _RE_BUILTIN_CONST
            sprobe("宿主枚举常量不进池（xlU| 不冒 xlUp）", "Sheet1", 6,
                "    xlU", 8, null, null, new[] { "xlUp" });
            // v155：宿主全局对象（Selection / ActiveSheet）不进池
            sprobe("宿主对象名不进池（Sele| 不冒 Selection）", "Sheet1", 6,
                "    Sele", 9, null, null, new[] { "Selection" });

            // ---- 2c. 隐式变量的可见范围 / Type 成员 ----
            sprobe("隐式变量不跨模块泄露（Sheet1 里 sec| 不冒模块2 的 secretVar）", "Sheet1", 6,
                "    sec", 8, null, null, new[] { "secretVar" });
            sprobe("本模块的隐式变量照常提示（模块2 里 sec|）", "模块2", 3,
                "    sec", 8, true, new[] { "secretVar" }, null);
            sprobe("Type 成员不进裸名池（模块1 里 Nic| 不冒 Nickname）", "模块1", 11,
                "    Nic", 8, null, null, new[] { "Nickname" });

            // ---- 2e. v183：窗体控件名（用户报"我在窗体里，不提示 TextBox1"）----
            //   控件名只活在**设计器数据**里，窗体代码模块里通常一个字都没有 ⇒ 只能由
            //   后端走 Designer 灌进 ModuleScope.Controls。这一组验可见性口径：
            //   在它自己的窗体里裸名可用，换个模块打同样的前缀不许冒出来。
            sprobe("窗体里打 te 出控件名（UserForm1）", "UserForm1", 3,
                "    s = te", 11, true, new[] { "TextBox1" }, null);
            sprobe("控件名不跨模块泄露（模块1 里 te|）", "模块1", 11,
                "    te", 7, null, null, new[] { "TextBox1" });
            // 一行代码都没有的新窗体：读代码时不能因为行数 0 就把组件整个跳过
            sprobe("空窗体也认控件（UserForm2 里 Li|）", "UserForm2", 1,
                "    Li", 7, true, new[] { "ListBox1" }, null);

            // ---- 2d. v158：用户案例（Set 出来的隐式变量必须能提示）----
            var dbg = Scope.Analyze("模块3", modUser);
            sb.AppendLine("  [诊断] Decls: " + string.Join(",", dbg.Decls.Select(d => d.Name + "@" + (d.Proc ?? "-"))));
            sb.AppendLine("  [诊断] FreeIdents: " + string.Join(",", dbg.FreeIdents));
            sb.AppendLine("  [诊断] Procs: " + string.Join(",", dbg.Procs.Select(p => p.Name + ":" + p.StartLine + "-" + p.EndLine)));
            sprobe("Set 出来的隐式变量能提示（condi| -> conditions）", "模块3", 7,
                "        If isMeet(condi", 24, true, new[] { "conditions" }, null);
            sprobe("同函数内 Set 出来的 dic 能提示（di| -> dic）", "模块4", 17,
                "    set getTitle2=di", 20, true, new[] { "dic" }, null);

            // ---- 3. A/B 对照：去掉 Option Explicit，同一个探针就该收进来 ----
            //   证明"未声明的不提示"确实是 Option Explicit 在起作用，不是别的原因。
            var noExpStd = modStd.Replace("Option Explicit", "              ");
            var synthNoExp = new SyntheticBackend()
                .Add("模块1", 1, noExpStd)
                .Add("cls1", 2, modCls);
            synthNoExp.Current = "模块1";
            var engNoExp = new Engine { Recent = new Recent(), Backend = synthNoExp };
            bool shown2 = engNoExp.Trigger(false,
                new Snapshot { Line = 11, Text = "    stud", Module = "模块1", CaretCol = 9 });
            bool leaked = shown2 && engNoExp.Items.Any(
                x => string.Equals(x.Name, "student", StringComparison.OrdinalIgnoreCase));
            if (!leaked) ok = false;
            sb.AppendLine((leaked ? "  PASS  " : "  FAIL  ") +
                          "对照：没有 Option Explicit 时 student 会被收进来");
            sb.AppendLine("         -> " +
                          (shown2 ? string.Join(",", engNoExp.Items.Take(8).Select(x => x.Name)) : "(不弹)"));
            if (shown2) engNoExp.Hide();

            // ---- 4. `New` 后面不弹（用户 v137 第 2 条，跟 `As` 一样）----
            sprobe("New 后不弹（dic.Add 1, New |）", "模块1", 11,
                "    dic.Add 1, New ", 21, false, null, null);
            sprobe("New 后半截词也不弹（New cl|）", "模块1", 11,
                "    dic.Add 1, New cl", 22, false, null, null);
            // ★这两条必须用"半截词"：`New |` / `As |`（词还没开始）本来就走
            //   "空词不弹"那条路，退掉 AfterTypeKeyword 也照样绿 —— 抓不到问题。
            //   真正会弹出全工程候选的是 `New cl|` / `As Lo|` 这种打了半个字的。
            sprobe("As 后不弹（For i As Lo|，行首不是 Dim）", "模块1", 11,
                "    For i As Lo", 17, false, null, null);

            // ---- 5. 类模块的 Public 成员不许混进别的模块的裸名池（用户 v137 第 3 条）----
            sprobe("类模块成员不进裸名池（模块1 里打 s|）", "模块1", 11,
                "    s", 6, null, null, new[] { "s1", "s2" });
            // ★v146 用户报："输入 cell 为什么会提示 cells，是不是泄露？"
            //   modMem 里只有 `wsR.Cells(1, 1).Value` —— 全是成员访问。
            //   在**另一个模块**里打 `cell` 时，候选里不许出现 Cells。
            sprobe("成员名不进裸名池（打 cell 不冒 Cells）", "模块1", 11,
                "    cell", 8, null, null, new[] { "Cells" });
            // ★v154：wsR 是 modMem 里"用过但没声明"的名字 —— VBA 会把它当成**该模块
            //   私有**的隐式变量。所以在 modMem 自己里必须提示，在别的模块里不该提示
            //   （那里写出来只会新建一个 Empty 变量，运行时报错）。
            sprobe("同一个词继续：在 modMem 自己里打 ws 能提示对象名 wsR", "modMem", 2,
                "    ws", 6, true, new[] { "wsR" }, null);
            sprobe("隐式变量不跨模块：模块1 里打 ws 不冒 wsR", "模块1", 11,
                "    ws", 6, null, null, new[] { "wsR" });
            // 对照：在 cls1 自己里面它就是裸名 —— 必须提示（证明按模块过滤，不是一刀切）
            sprobe("对照：cls1 自己里 s1 照常提示（s|）", "cls1", 4,
                "s", 2, true, new[] { "s1" }, null);
            // ---- 6. 点号之后一整个让给 VBE（v138，与 Python 版 v103c 同口径）----
            // 默认（MEMBER_POPUP=0）一个候选都不出：VBE 在这个位置按真类型弹真成员。
            sprobe("点号后不弹（c.|）", "模块1", 11,
                "    c.", 7, false, null, null);
            sprobe("点号后不弹（c.s|，成员名打到一半）", "模块1", 11,
                "    c.s", 8, false, null, null);
            sprobe(") 之后不弹（Range(\"A1\").Fo|）", "模块1", 11,
                "    Range(\"A1\").Fo", 19, false, null, null);
            // 词要用能匹配到东西的（`.Se` 会命中 Set / Select / …）——
            // 用 `.Size` 这种"恰好一个候选都匹配不上"的词，退掉判据也照样绿，
            // 等于没验（v137 的 A/B 已经栽过一次这种"探针空转"）。
            sprobe("With 块点号不弹（    .Se|）", "模块1", 11,
                "    .Se", 9, false, null, null);
            {
                // ---- A/B：MEMBER_POPUP=1 时同一位置必须重新弹（且按类型收窄）----
                // 这一条是"点号后不弹"的**反证**：退掉那条判据它就绿、留着它才红不了；
                // 反过来一旦判据写错（比如把 `c.` 判成非成员位置），这条就红。
                synth.Current = "模块1";
                var eOn = new Engine
                {
                    Recent = new Recent(),
                    Backend = synth,
                    MemberPopup = true
                };
                eOn.InvalidatePool();
                bool shownOn = eOn.Trigger(false,
                    new Snapshot { Line = 11, Text = "    c.", Module = "模块1", CaretCol = 7 });
                bool hasMembers = shownOn && eOn.Items.Any(
                    x => string.Equals(x.Name, "s1", StringComparison.OrdinalIgnoreCase));
                if (!(shownOn && hasMembers)) ok = false;
                sb.AppendLine(((shownOn && hasMembers) ? "  PASS  " : "  FAIL  ") +
                              "A/B：MEMBER_POPUP=1 点号后照弹，按类型收窄到 s1/s2");
                sb.AppendLine("         -> " +
                    (shownOn ? string.Join(",", eOn.Items.Take(8).Select(x => x.Name)) : "(不弹)"));
                if (shownOn) eOn.Hide();
            }

            sb.AppendLine(ok ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "enginetest.log"), sb.ToString(), Encoding.UTF8);
            Console.WriteLine(sb.ToString());
            return ok ? 0 : 1;
        }

        /// <summary>
        /// 完全合成的工程（不碰 COM）：模块名 -> 全文 + 组件类型。
        /// 用来钉"跨模块可见性"这类与当前打开哪个模块无关的契约 —— 真机探针在没开代码窗时
        /// 一律 skip，等于什么都没验（"绿着但没测"就是这么来的）。
        /// </summary>
        private sealed class SyntheticBackend : IVbeSource
        {
            private readonly Dictionary<string, string> _texts =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, int> _types =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            // 窗体控件名（ComponentType=3 才有）。真机上由 VbeBackend 走 Designer 读，
            // 合成后端直接摆 —— 这一格验的是"控件名有没有进池 / 该不该跨模块可见"。
            private readonly Dictionary<string, List<string>> _ctrls =
                new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            public string Current;

            public SyntheticBackend Add(string name, int type, string text)
            {
                _texts[name] = text;
                _types[name] = type;
                return this;
            }

            /// <summary>加一个窗体（ComponentType=3）并给它挂上控件名。</summary>
            public SyntheticBackend AddForm(string name, string text, params string[] controls)
            {
                Add(name, 3, text);
                _ctrls[name] = new List<string>(controls ?? Array.Empty<string>());
                return this;
            }

            public Snapshot Snapshot(bool verify = true) => null;

            public string ModuleText() =>
                _texts.TryGetValue(Current ?? "", out var t) ? t : "";

            public List<string> ComponentNames() => new List<string>(_texts.Keys);

            public List<ModuleScope> ProjectScopes(out int signature)
            {
                signature = 1;
                var list = new List<ModuleScope>();
                foreach (var kv in _texts)
                {
                    var ms = Scope.Analyze(kv.Key, kv.Value);
                    ms.ComponentType = _types[kv.Key];
                    List<string> cs;
                    if (_ctrls.TryGetValue(kv.Key, out cs)) ms.Controls = cs;
                    list.Add(ms);
                }
                return list;
            }

            public bool ApplyCompletion(int lineNo, int wordStartCol, int endCol,
                string completion, bool wantParens, bool wantSpace, string enclosingProc) => false;

            public string EnclosingProcName() => "test";
        }
    }
}
