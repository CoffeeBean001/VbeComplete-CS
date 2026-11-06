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
    /// <summary>真机键探针：括号补全、钩子、回车结构补全。</summary>
    internal static partial class Program
    {
        /// <summary>
        /// 括号补全验证（真机）：走 Engine.AcceptName 这条真路径，确认选函数真的补一对括号、
        /// 光标落在括号中间，类型名 / 关键字一个都不许补。全程只动模块末尾临时行，最后还原。
        /// ★括号逻辑分散在四处（判档位 → 判位置 → 判 Call 语句 → 落光标），只有真机能验完整条。
        /// </summary>
        private static int ParenTest()
        {
            var sb = new StringBuilder();
            var be = new VbeBackend();
            string orig = be.ModuleText();
            int count = be.LineCount();
            bool ok = true;

            sb.AppendLine("=== paren test (真机) ===");
            sb.AppendLine("module lines = " + count);
            if (count <= 0)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "paretest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }

            var eng = new Engine { Backend = be };

            // 档位判据先单独对一遍（不碰 COM）
            bool wp, ws;
            eng.Followup("MsgBox", null, out wp, out ws);
            sb.AppendLine("Followup(MsgBox) -> parens=" + wp + " space=" + ws);
            ok &= wp;
            eng.Followup("String", null, out wp, out ws);
            sb.AppendLine("Followup(String 类型) -> parens=" + wp + " space=" + ws);
            ok &= !wp;
            eng.Followup("Dim", null, out wp, out ws);
            sb.AppendLine("Followup(Dim 关键字) -> parens=" + wp + " space=" + ws);
            ok &= !wp && ws;

            // (探针行, 光标列, 候选名, 期望行, 期望光标列)
            var cases = new[]
            {
                // ★用户报的那条：一行的开头（只有缩进）
                ("    MsgB", 9, "MsgBox", "    MsgBox()", 12),
                // 赋值右侧
                ("    x = Spl", 13, "Split", "    x = Split()", 15),
                // 类型名：绝不能补括号
                ("    Dim y As Str", 17, "String", "    Dim y As String", 20),
                // ★v152（对标 IDEA）：光标停在词**中间**时，替换的是整个标识符
                //   —— 只换光标左边那一截会得到 `    MsgBoxx`。
                ("    MsgBx", 9, "MsgBox", "    MsgBox()", 12),
            };

            try
            {
                foreach (var c in cases)
                {
                    be.InsertLines(count + 1, c.Item1);
                    be.SetCaret(count + 1, c.Item2);
                    bool acc = eng.AcceptName(c.Item3);
                    string line = be.LineAt(count + 1) ?? "";
                    var snap = be.Snapshot();
                    int caret = snap == null ? -1 : snap.CaretCol;

                    bool lineOk = line == c.Item4;
                    bool caretOk = caret == c.Item5;
                    if (!acc || !lineOk || !caretOk) ok = false;
                    sb.AppendLine("  [" + (acc && lineOk && caretOk ? "PASS" : "FAIL") + "] "
                        + c.Item1.TrimEnd() + " + Tab " + c.Item3);
                    sb.AppendLine("         行 = [" + line + "]  want [" + c.Item4 + "]"
                        + (lineOk ? "" : "  <- 行不对"));
                    sb.AppendLine("         光标 = " + caret + "  want " + c.Item5
                        + (caretOk ? "" : "  <- 落点不对"));

                    // 立刻删掉这一行，下一个用例从同一个基准行数开始
                    for (int i = 0; i < 3; i++)
                    {
                        int c2 = be.LineCount();
                        if (c2 <= count) break;
                        be.DeleteLines(count + 1, c2 - count);
                    }
                }
            }
            finally
            {
                for (int i = 0; i < 3; i++)
                {
                    int c2 = be.LineCount();
                    if (c2 <= count) break;
                    be.DeleteLines(count + 1, c2 - count);
                }
            }

            string now = be.ModuleText();
            bool restored = now == orig;
            sb.AppendLine("restored = " + restored);
            if (!restored)
            {
                ok = false;
                int n = Math.Min(now == null ? 0 : now.Length, orig == null ? 0 : orig.Length);
                int at = 0;
                while (at < n && now[at] == orig[at]) at++;
                sb.AppendLine("首个差异位置 = " + at + "（len now=" + (now == null ? -1 : now.Length)
                    + " orig=" + (orig == null ? -1 : orig.Length) + "）");
                int from = Math.Max(0, at - 40);
                sb.AppendLine("orig[" + from + "..] = [" + SafeSlice(orig, from, 120) + "]");
                sb.AppendLine("now [" + from + "..] = [" + SafeSlice(now, from, 120) + "]");
            }
            sb.AppendLine(ok && restored ? "RESULT: PASS" : "RESULT: FAIL");

            File.WriteAllText(Path.Combine(Settings.Home, "paretest.log"), sb.ToString(), Encoding.UTF8);
            return ok && restored ? 0 : 1;
        }

        /// <summary>
        /// 真机验证四个命令：Ctrl+Shift+Enter / Ctrl+D / Ctrl+K / Ctrl+/。
        /// ★按键能不能走到钩子在这里试不了（钩子按设计不理注入键），但"命令本身能不能干活"
        ///   可以 100% 验 —— 这里调的就是钩子调的那几个 Commands.*。全过 ⇒ 问题只在路由。
        /// 全程只动模块末尾临时追加的 6 行，最后删干净并逐字节比对原文。
        /// </summary>
        private static int HkTest()
        {
            var sb = new StringBuilder();
            var be = new VbeBackend();
            string orig = be.ModuleText();
            int baseCount = be.LineCount();
            bool allOk = true;

            sb.AppendLine("=== hotkey test (真机) ===");
            sb.AppendLine("module lines = " + baseCount + "  len = " + (orig == null ? -1 : orig.Length));
            if (baseCount <= 0)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "hktest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }

            var block = new[]
            {
                "Sub VbcHkTest()",
                "    Dim n As Long",
                "    n = 1",
                "    n = 2",
                "",
                "End Sub",
            };
            for (int i = 0; i < block.Length; i++)
                be.InsertLines(baseCount + 1 + i, block[i]);

            // ★v149：`Ctrl+Shift+↑/↓` 那组用例随功能一起删了，Lblank / Lhead / Lend
            //   随之失去用途（原本就只被那两条用例引用）。
            int Ldim = baseCount + 2, L1 = baseCount + 3, L2 = baseCount + 4;
            string dimText = be.LineAt(Ldim) ?? "";
            string oneText = be.LineAt(L1) ?? "";
            string twoText = be.LineAt(L2) ?? "";
            sb.AppendLine("scratch: dim=[" + dimText + "] one=[" + oneText + "] two=[" + twoText + "]");

            Action<string, bool, string> check = (name, ok, detail) =>
            {
                if (!ok) allOk = false;
                sb.AppendLine((ok ? "  PASS  " : "  FAIL  ") + name + "   " + detail);
            };

            try
            {
                // ---- 0. SetCaret 的列语义（五个命令全都踩在它上面）----
                {
                    be.SetCaret(L1, 1);
                    be.InvalidateSnapshot();
                    int c1 = be.Snapshot()?.CaretCol ?? -1;
                    be.SetCaret(L1, 3);
                    be.InvalidateSnapshot();
                    int c3 = be.Snapshot()?.CaretCol ?? -1;
                    check("SetCaret(.,1)/(.,3) 列语义", c1 == 1 && c3 == 3,
                        "读回 col=" + c1 + "/" + c3 + "  want 1/3");
                }

                // ---- 1. Ctrl+D 复制行 ----
                {
                    be.SetCaret(L1, 1);
                    be.InvalidateSnapshot();
                    int before = be.LineCount();
                    bool r = Commands.DuplicateLine(be);
                    int after = be.LineCount();
                    var snap = be.Snapshot();
                    string a = be.LineAt(L1) ?? "", b = be.LineAt(L1 + 1) ?? "";
                    bool ok = r && after == before + 1
                              && Rules.Rstrip(a) == Rules.Rstrip(b)
                              && Rules.Rstrip(b).Trim() == "n = 1"
                              && snap != null && snap.Line == L1 + 1;
                    check("Ctrl+D 复制行", ok,
                        "ret=" + r + " lines " + before + "->" + after +
                        " dup=[" + b + "] caretLine=" + (snap?.Line ?? -1));
                    // 复原：副本还在 L1+1
                }

                // ---- 2. Ctrl+K 删行（光标落到上一行行尾）----
                {
                    int before = be.LineCount();
                    be.SetCaret(L1 + 1, 1);           // 刚复制出来的那一行
                    be.InvalidateSnapshot();
                    bool r = Commands.DeleteLine(be);
                    int after = be.LineCount();
                    var snap = be.Snapshot();
                    string now = be.LineAt(L1 + 1) ?? "";
                    bool ok = r && after == before - 1
                              && Rules.Rstrip(now) == Rules.Rstrip(twoText)
                              && snap != null && snap.Line == L1;
                    check("Ctrl+K 删行落到上一行", ok,
                        "ret=" + r + " lines " + before + "->" + after +
                        " next=[" + now + "] caretLine=" + (snap?.Line ?? -1) + " want=" + L1);
                }

                // ---- 3. Ctrl+/ 注释开关 ----
                {
                    be.SetCaret(L1, 1);
                    be.InvalidateSnapshot();
                    var s0 = be.Snapshot();
                    int sl, sc, el, ec;
                    bool hasSel = be.TryGetSelection(out sl, out sc, out el, out ec);
                    bool r1 = Commands.ToggleComment(be);
                    string c1 = be.LineAt(L1) ?? "";
                    be.InvalidateSnapshot();
                    bool r2 = Commands.ToggleComment(be);
                    string c2 = be.LineAt(L1) ?? "";
                    bool ok = r1 && r2
                              && c1.TrimStart().StartsWith("'", StringComparison.Ordinal)
                              && c1.Contains("n = 1")
                              && Rules.Rstrip(c2) == Rules.Rstrip(oneText);
                    check("Ctrl+/ 注释开关", ok,
                        "ret=" + r1 + "/" + r2 + " 加后=[" + c1 + "] 还原后=[" + c2 + "]");
                    sb.AppendLine("        snapLine=" + (s0?.Line ?? -1) + " sel=" + hasSel +
                                  "(" + sl + "," + sc + "," + el + "," + ec + ")" +
                                  " 原行=[" + oneText + "] 期望还原=[" + Rules.Rstrip(oneText) + "]");
                }

                // ---- 4. Ctrl+Shift+Enter 上方插行 ----
                {
                    int before = be.LineCount();
                    be.SetCaret(L1, 1);
                    be.InvalidateSnapshot();
                    bool r = Commands.InsertLineAbove(be);
                    int after = be.LineCount();
                    string ins = be.LineAt(L1) ?? "";
                    string below = be.LineAt(L1 + 1) ?? "";
                    int indent = Commands.LeadingWs(oneText).Length;
                    bool ok = r && after == before + 1
                              && Rules.Rstrip(below) == Rules.Rstrip(oneText)
                              && Commands.LeadingWs(ins).Length == indent;
                    check("Ctrl+Shift+Enter 上方插行", ok,
                        "ret=" + r + " lines " + before + "->" + after +
                        " ins=[" + ins + "] 缩进=" + Commands.LeadingWs(ins).Length + " want=" + indent +
                        " below=[" + below + "] want=[" + oneText + "]");
                    // ★无论成败都要把这行临时行删掉 —— 否则后面几条用例的行号全偏，
                    //   一个失败会伪装成三个失败（第一次跑就是这样）。
                    if (after > before)
                    {
                        be.SetCaret(L1, 1);
                        be.InvalidateSnapshot();
                        Commands.DeleteLine(be);
                    }
                }

                // ==========================================================
                // 以下六条是"编辑手感"命令：每条自己造一小段临时代码、验完立刻整段删掉，
                // 不依赖上面那 6 行的行号。
                // ⚠️ 功能删了用例必须一起删（留着编译不过）。
                // ==========================================================
                Func<string[], int> append = lines =>
                {
                    int at = be.LineCount() + 1;
                    // ★倒着插（全部插在 at 这一位）：正着插的话，`""`（空行）会在
                    //   某一瞬间成为**模块最后一行**，而 VBE 对"末尾的空行"处理
                    //   并不可靠（可能直接吞掉），后面几条用例的行号就全偏了。
                    for (int i = lines.Length - 1; i >= 0; i--) be.InsertLines(at, lines[i]);
                    return at;
                };
                Action dropFrom = () =>
                {
                    for (int i = 0; i < 3; i++)
                    {
                        int c = be.LineCount();
                        if (c <= baseCount) break;
                        be.DeleteLines(baseCount + 1, c - baseCount);
                    }
                };

                // ---- 7. Shift+Enter：下方插一行，缩进抄本行 ----
                {
                    int at = append(new[] { "Sub VbcNl()", "    n = 1", "End Sub" });
                    be.SetCaret(at + 1, 1);
                    be.InvalidateSnapshot();
                    bool r = be.NewLineBelow(false, false, false);
                    string ins = be.LineAt(at + 2) ?? "";
                    var snap = be.Snapshot();
                    bool ok = r && ins == "    " && Rules.Rstrip(be.LineAt(at + 1) ?? "") == "    n = 1"
                              && snap != null && snap.Line == at + 2;
                    check("Shift+Enter 下方插行", ok,
                        "ret=" + r + " 新行=[" + ins.Replace("\t", "\\t") + "] 期望=[    ] caret=" + (snap?.Line ?? -1));
                    dropFrom();
                }

                // ---- 8. 回车自动缩进 + 补收尾 ----
                {
                    int at = append(new[] { "Sub VbcIn()", "    For i = 1 To 3", "End Sub" });
                    // ★行尾 = "字符数 + 1"（SetCaret 的列是"第几个字符之前"）
                    be.SetCaret(at + 1, (be.LineAt(at + 1) ?? "").Length + 1);
                    be.InvalidateSnapshot();
                    bool r = be.NewLineBelow(true, true, true);
                    string line2 = be.LineAt(at + 2) ?? "";
                    string line3 = be.LineAt(at + 3) ?? "";
                    var snap = be.Snapshot();
                    bool ok = r && line2 == "        " && Rules.Rstrip(line3) == "    Next"
                              && snap != null && snap.Line == at + 2;
                    check("回车自动缩进 + 补 Next", ok,
                        "ret=" + r + " 新行=[" + line2 + "] 下一行=[" + line3 + "] caret=" + (snap?.Line ?? -1));
                    dropFrom();
                }

                // ---- 9. 退格删整行空白 ----
                {
                    int at = append(new[] { "Sub VbcBs()", "    ", "    n = 2", "End Sub" });
                    string raw = be.LineAt(at + 1) ?? "";
                    be.SetCaret(at + 1, raw.Length + 1);
                    be.InvalidateSnapshot();
                    int before = be.LineCount();
                    bool r = be.DeleteBlankLineHere();
                    int after = be.LineCount();
                    if (raw.Length == 0)
                    {
                        // VBE 把只有空白的行 trim 成真空行了 —— 那时"原生退格"的结果
                        // 与删整行一样，我们不接管是**正确**的
                        check("退格删整行空白（跳过：VBE 已 trim 成真空行）", !r,
                            "raw 为空 -> 不接管，ret=" + r);
                    }
                    else
                    {
                        check("退格删整行空白", r && after == before - 1 && (be.LineAt(at + 1) ?? "").Trim() == "n = 2",
                            "ret=" + r + " lines " + before + "->" + after + " 下一行=[" + be.LineAt(at + 1) + "]");
                    }
                    dropFrom();
                }

                // ---- 10. 退格删一对空括号 ----
                {
                    int at = append(new[] { "Sub VbcDp()", "    x = f()", "End Sub" });
                    be.SetCaret(at + 1, 11);               // 夹在 `(` 与 `)` 中间
                    be.InvalidateSnapshot();
                    bool r = be.DeletePairHere();
                    string now = be.LineAt(at + 1) ?? "";
                    var snap = be.Snapshot();
                    bool ok = r && string.Equals(Rules.Rstrip(now), "    x = f",
                        StringComparison.OrdinalIgnoreCase) && snap != null && snap.CaretCol == 10;
                    check("退格删一对空括号", ok,
                        "ret=" + r + " 行=[" + now + "] caretCol=" + (snap?.CaretCol ?? -1) + " want=10");
                    // ★大小写不敏感：VBE 会把单字母标识符按工程里同名符号改大小写
                    //   （这份工作簿里 `f` 落盘成 `F`），按 Ordinal 比就是一条假红
                    //   —— 和 --enterprobe 里 WPS 那条 `x = 1` / `X = 1` 同一个坑。
                    dropFrom();
                }

                // ---- 11. 自动配对括号 ----
                {
                    int at = append(new[] { "Sub VbcAu()", "    abc" });
                    be.SetCaret(at + 1, (be.LineAt(at + 1) ?? "").Length + 1);   // 行尾
                    be.InvalidateSnapshot();
                    bool r = be.InsertPair('(');
                    string now = be.LineAt(at + 1) ?? "";
                    var snap = be.Snapshot();
                    int col = snap == null ? -1 : snap.CaretCol;
                    // "    abc()" 共 9 个字符，光标落在 `(` 与 `)` **之间** = 第 9 列
                    // （第 k 列 = "第 k 个字符之前"，`)` 是第 9 个字符）
                    bool ok = r && Rules.Rstrip(now) == "    abc()" && col == 9;
                    check("自动配对括号", ok,
                        "ret=" + r + " 行=[" + now + "] caretCol=" + col + " want=9");
                    dropFrom();
                }

                // ---- 12. 插入代码片段（多行） ----
                {
                    // ★中间留一行**空行**当落点：片段的落点模型是"整行拼接"
                    //   （前半行 + 片段首行 / 片段末行 + 后半行），光标停在有内容的
                    //   行尾时会拼成 `    x = 0    a = 1` —— 那是设计如此，不是 bug。
                    //   拿空行当落点才是用户真实用法：点一下空行，插入。
                    int at = append(new[] { "Sub VbcSn()", "", "End Sub" });
                    be.SetCaret(at + 1, 1);
                    be.InvalidateSnapshot();
                    int before = be.LineCount();
                    bool r = be.InsertSnippet("    a = 1\r\n    b = 2");
                    int after = be.LineCount();
                    string l1 = be.LineAt(at + 1) ?? "";
                    string l2 = be.LineAt(at + 2) ?? "";
                    string l3 = be.LineAt(at + 3) ?? "";
                    bool ok = r && after == before + 1
                              && Rules.Rstrip(l1) == "    a = 1"
                              && Rules.Rstrip(l2) == "    b = 2"
                              && Rules.Rstrip(l3) == "End Sub";
                    check("插入多行片段", ok,
                        "ret=" + r + " lines " + before + "->" + after +
                        " [" + l1 + "][" + l2 + "][" + l3 + "]");
                    dropFrom();
                }

                // ---- 13. （原 Ctrl+Shift+L 格式化缩进）----
                // v150 用户要求删除该功能，用例与实现一起移除。编号留着不复用，
                // 免得跟历史日志 / 提交信息里的"第 13 条"对不上。

                // ---- 14. v141：光标跳到别的行（鼠标单击 / 方向键）不许弹窗 ----
                // 用户报的 bug："有时候鼠标单击到某个变量名，它会自动弹提示词框。"
                // 这条验的是那个判据**依赖的事实**：真机上"光标换行"确实会让
                // Snapshot.Line 变掉（判据靠它把鼠标操作和打字分开）。
                {
                    var blk = new[] { "Sub VbcJump()", "    n = 1", "    n = 2", "End Sub" };
                    int at = append(blk);
                    be.SetCaret(at + 1, 1);                       // 模拟"刚才在这一行打字"
                    be.InvalidateSnapshot();
                    var a = be.Snapshot();
                    // 模拟鼠标单击下一行那个变量名的词尾
                    be.SetCaret(at + 2, (be.LineAt(at + 2) ?? "").Length + 1);
                    be.InvalidateSnapshot();
                    var b = be.Snapshot();
                    bool jumped = a != null && b != null && b.Line != a.Line;
                    bool want = Rules.AutoTriggerWanted(true, false, true, jumped);
                    check("跳行不弹（鼠标单击 / 方向键）", jumped && !want,
                        "L" + (a?.Line ?? -1) + "->L" + (b?.Line ?? -1) +
                        " jumped=" + jumped + " want=" + want);
                    dropFrom();
                }
            }
            finally
            {
                // 删干净临时块（多删几次兜底），再逐字节比对原文
                for (int i = 0; i < 4; i++)
                {
                    int c = be.LineCount();
                    if (c <= baseCount) break;
                    be.DeleteLines(baseCount + 1, c - baseCount);
                }
            }

            string now2 = be.ModuleText();
            bool restored = string.Equals(now2, orig, StringComparison.Ordinal);
            if (!restored) allOk = false;
            sb.AppendLine("restored = " + restored + "   lines " + be.LineCount() + "/" + baseCount);
            sb.AppendLine(allOk ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "hktest.log"), sb.ToString(), Encoding.UTF8);
            return allOk ? 0 : 1;
        }

        /// <summary>
        /// 真机"回车补结构"探针 —— **走真键盘钩子**（真 App + 真轮询 + SendInput）。
        /// ★--hktest 只验后端（直接调 NewLineBelow），而用户走的是 WH_KEYBOARD_LL →
        ///   FreshHookSnap → EnterIndentWanted → Post → NewLineBelow。后端全绿而真机不补
        ///   的情况真出现过，所以必须验"手指走的那条路"。
        /// </summary>
        private static int EnterProbe()
        {
            var sb = new StringBuilder();
            var gate = new object();
            bool allOk = true;
            Action<string, bool, string> check = (name, ok, detail) =>
            {
                if (!ok) allOk = false;
                lock (gate) sb.AppendLine((ok ? "  PASS  " : "  FAIL  ") + name + "   " + detail);
            };

            sb.AppendLine("=== enter probe（真机 / 走真钩子） ===");
            sb.AppendLine("（回车补结构 与 Ctrl+/ 注释都在这条真路径上验：钩子 -> 判据 -> Post -> COM）");

            var app = new App();
            var be = new VbeBackend();          // 主线程（STA）自己持有，不用 RunOnUi 绕
            int leftover = CleanupScratch(be);
            string orig = be.ModuleText();
            int baseCount = be.LineCount();

            sb.AppendLine("module lines = " + baseCount + "  len = " + (orig == null ? -1 : orig.Length) +
                          "  上次残留块清理 = " + leftover);
            // baseCount 可以是 0（空模块）：探针往里插脚手架再删干净，正好也能验
            if (baseCount < 0 || orig == null)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "enterprobe.log"), sb.ToString(), new UTF8Encoding(false));
                return 1;
            }

            // 造一段固定的脚手架：一个 Sub + 一行现成的 For + End Sub
            int at = baseCount + 1;
            RebuildScaffold(app, be, at, "    For i = 1 To 3");
            be.FocusVbe();

            // ★Start() 必须显式调用：装钩子 / 起托盘 / 起轮询定时器都在它里面
            app.Start();

            string diagA = "", diagA2 = "", diagB = "", diagC = "";
            int imeModeAtStart = -1;   // 探针把输入法切到英数档前记下原模式，跑完还回去

            var driver = new Thread(() =>
            {
                try
                {
                    // 1. 等轮询把"焦点在代码窗"跑出来 —— 钩子的第一道闸就是它。
                    //    ★从 bash / 计划任务起的进程 SetForegroundWindow 会被系统拒掉，
                    //      所以每轮都硬抢一次（AttachThreadInput 那套），抢到为止。
                    bool focused = false;
                    for (int i = 0; i < 40 && !focused; i++)
                    {
                        app.RunOnUi(() =>
                        {
                            Native.ForceForeground(Native.FindWindow("wndclass_desked_gsk", null));
                            be.SetCaret(at + 1, 19);
                            be.InvalidateSnapshot();
                        });
                        Thread.Sleep(250);
                        focused = app.CodePaneFocused;
                    }
                    check("焦点在代码窗（钩子的第一道闸）", focused, "CodePaneFocused=" + focused);
                    if (!focused) return;

                    // 输入法预检：B/C 是"逐字符打字"的用例，中文模式下每个字母都会被
                    // 输入法拿去组字（实测敲出来是 `Fori= 1 To3` / `If下》 0 Then`），
                    // 那不是程序坏了，是环境不对。这里直接把它切到英数档（走
                    // WM_IME_CONTROL），跑完再恢复成用户原来的模式。
                    int imeMode0 = Native.ImeConversionMode(Native.FocusedVbeWindow());
                    bool alnum0 = Native.ImeAlphanumeric(Native.FocusedVbeWindow());
                    if (!alnum0)
                    {
                        Native.ImeSetAlphanumeric(Native.FocusedVbeWindow());
                        Thread.Sleep(200);
                    }
                    bool alnum1 = Native.ImeAlphanumeric(Native.FocusedVbeWindow());
                    imeModeAtStart = imeMode0;
                    lock (gate)
                        sb.AppendLine("IME 英数模式 = " + alnum0 + " -> " + alnum1 +
                            "（conv=0x" + imeMode0.ToString("X2") + "，跑完恢复）" +
                            (alnum1 ? "" : "   ⚠️ 切不过去：B/C 会被输入法吃掉"));

                    // 2. 用例 A：现成的行，下面紧跟 End Sub，直接回车
                    Thread.Sleep(300);
                    diagA = EnterDiag(app);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(700);
                    string[] a = ReadLines(app, be, at, 4);
                    check("A 现成 For 行 + 回车 -> 补 Next", a[3] == "    Next",
                        Tail(a) + "  diag{" + diagA + "}");

                    // 3. 用例 A2（★用户报的那一条）：过程中间插一个 For，
                    //    下面压着同级的旧语句 —— Python 版会补，C# 旧口径不补。
                    RebuildScaffold(app, be, at, "    For i = 1 To 3", "    x = 1", "    y = 2");
                    Thread.Sleep(300);
                    diagA2 = EnterDiag(app);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(700);
                    string[] a2 = ReadLines(app, be, at, 6);
                    check("A2 同级旧代码压着也要补 Next（用户报的那一条）",
                        // ★大小写不敏感：WPS 的 VBE 会把 `x = 1` 自动改成 `X = 1`（自动大写），
                        //   拿 Ordinal 比会把它判成失败 —— 那不是回归，是宿主编辑器行为。
                        a2[3] == "    Next" &&
                        string.Equals(a2[4], "    x = 1", StringComparison.OrdinalIgnoreCase),
                        "[" + a2[1] + "][" + a2[2] + "][" + a2[3] + "][" + a2[4] + "]  diag{" + diagA2 + "}");

                    // 4. 用例 B：空行上逐字符敲 For 行（模拟真人），最后一下立刻回车
                    RebuildScaffold(app, be, at, "");
                    Thread.Sleep(300);
                    bool typedB = TypeTyping(app, "    For i = 1 To 3", 45, out string howB);
                    diagB = EnterDiag(app);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(700);
                    string[] b = ReadLines(app, be, at, 4);
                    check("B 逐字符敲 For + 立刻回车 -> 补 Next", typedB && b[3] == "    Next",
                        "typed=" + typedB + " via=" + howB + " " + Tail(b) + "  diag{" + diagB + "}");

                    // 5. 用例 C：If … Then -> End If
                    RebuildScaffold(app, be, at, "");
                    Thread.Sleep(300);
                    bool typedC = TypeTyping(app, "    If x > 0 Then", 45, out string howC);
                    diagC = EnterDiag(app);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(700);
                    string[] c = ReadLines(app, be, at, 4);
                    check("C 逐字符敲 If … Then -> 补 End If", typedC && c[3] == "    End If",
                        "typed=" + typedC + " via=" + howC + " " + Tail(c) + "  diag{" + diagC + "}");

                    // 5b. 用例 C2 / C3（v184，用户报的）：`If … Then` 之后另起一行写
                    //     `ElseIf … Then`，再按回车 —— 既要把 ElseIf 拉回 `If` 的层级，
                    //     也要把新行缩进一级。两条分别验**后端**与**钩子那条真路**：
                    //       C2 静态摆好（快照肯定完整）—— 红的说明 BranchAlign/后端坏了；
                    //       C3 逐字符敲完立刻回车（快照常落后一格）—— 红的说明钩子判据
                    //          没把 `ElseIf` 认成"该接管的行"，回车直接漏给 VBE 原生。
                    int c2col = 0;
                    RebuildScaffold(app, be, at, "    If x > 0 Then", "        ElseIf x < 0 Then", "    End If");
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.SetCaret(at + 2, "        ElseIf x < 0 Then".Length + 1);
                        be.InvalidateSnapshot();
                    });
                    Thread.Sleep(300);
                    string diagC2 = EnterDiag(app);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(700);
                    string[] c2 = ReadLines(app, be, at, 5);
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.TryGetSelection(out int _, out int col, out int _, out int _);
                        c2col = col;
                    });
                    check("C2 ElseIf 行回车（静态摆放）：拉回 If 层级 + 新行缩进一级",
                        string.Equals(c2[2], "    ElseIf x < 0 Then", StringComparison.OrdinalIgnoreCase)
                        && c2[3].Trim() == "" && c2col == 9 && c2col > 0
                        && string.Equals(c2[4], "    End If", StringComparison.OrdinalIgnoreCase),
                        Tail(c2) + " col=" + c2col + "（want 「    ElseIf x < 0 Then」/8空格/col9）  diag{" + diagC2 + "}");

                    // 5c. 用例 C3 续：**必须制造出"快照落后一格"才算验到**。
                    //     上面 C3 那一轮真机上常常是 want=True（轮询恰好赶上了），那就只是
                    //     C2 的重复 —— 用户报的那条恰恰发生在"最后一个字符刚落下、100ms 内
                    //     就回车"的时刻。这里把最后一个 `n` 的延迟设为 0：敲完立刻回车，
                    //     快照里那一行必然还差一格。三轮都造不出来就判 FAIL 并明说
                    //     "这条没验到"，绝不拿假前提刷绿。
                    int c3col = 0;
                    bool sawLag = false, c3ok = false, typedC3 = false;
                    string c3diag = "";
                    for (int attempt = 0; attempt < 3 && !sawLag; attempt++)
                    {
                        RebuildScaffold(app, be, at, "    If x > 0 Then", "        ", "    End If");
                        app.RunOnUi(() =>
                        {
                            be.ForceReloadModule();
                            be.SetCaret(at + 2, 9);
                            be.InvalidateSnapshot();
                        });
                        Thread.Sleep(250);
                        typedC3 = TypeTyping(app, "ElseIf x < 0 The", 45, out string howC3);
                        typedC3 &= TypeTyping(app, "n", 0, out _);
                        string diagC3 = EnterDiag(app);
                        SendKeys(app, Native.VK_RETURN);
                        Thread.Sleep(700);
                        string[] c3 = ReadLines(app, be, at, 5);
                        app.RunOnUi(() =>
                        {
                            be.ForceReloadModule();
                            be.TryGetSelection(out int _, out int col, out int _, out int _);
                            c3col = col;
                        });
                        c3ok = typedC3
                            && string.Equals(c3[2], "    ElseIf x < 0 Then", StringComparison.OrdinalIgnoreCase)
                            && c3[3].Trim() == "" && c3col == 9 && c3col > 0;
                        c3diag = "try" + (attempt + 1) + " via=" + howC3 + " " + Tail(c3)
                                 + " col=" + c3col + "  diag{" + diagC3 + "}";
                        sawLag = diagC3.Contains("want=False");
                    }
                    check("C3 快照落后一格时敲 ElseIf + 立刻回车：同样要拉齐 + 缩进一级",
                        sawLag && c3ok,
                        (sawLag ? "" : "⚠️ 三轮都没制造出快照落后 -> 这条没验到  ") + c3diag);

                    // 6. 用例 D：单选 If（`If x Then y = 1`）—— 不该补 End If，
                    //    但**回车必须照常换行**（这一档走的是"先接管、判不过再还回去"，
                    //    最容易写出"吞了却什么都没发生"的就是这条路）。
                    RebuildScaffold(app, be, at, "    If x Then y = 1");
                    int beforeD = 0;
                    app.RunOnUi(() => { be.ForceReloadModule(); beforeD = be.LineCount(); });
                    Thread.Sleep(300);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(800);
                    string[] d = ReadLines(app, be, at, 4);
                    int afterD = 0;
                    app.RunOnUi(() => { be.ForceReloadModule(); afterD = be.LineCount(); });
                    check("D 单选 If 回车：不补 End If 但必须换行",
                        afterD == beforeD + 1 &&
                        string.Equals(d[1], "    If x Then y = 1", StringComparison.OrdinalIgnoreCase) &&
                        d[2].Trim() == "",
                        "lines " + beforeD + "->" + afterD + " " + Tail(d));

                    // 7. 用例 E：Ctrl+/ 注释开关 —— 走【同一条真键盘路径】。
                    //    ★命令本身另有 --hktest 直接调 Commands 验过；这里要验的是
                    //      "钩子有没有把 Ctrl+/ 路由进那个分支、有没有把键吞了却没做事"。
                    //      用户报"Ctrl+/ 失效"时，只有这条路能证明卡在哪。
                    RebuildScaffold(app, be, at, "    n = 1");
                    Thread.Sleep(300);
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(700);
                    string[] e1 = ReadLines(app, be, at, 3);
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(700);
                    string[] e2 = ReadLines(app, be, at, 3);
                    check("E Ctrl+/ 整行注释开关（真钩子路径）",
                        e1[1].TrimStart().StartsWith("'", StringComparison.Ordinal)
                        && e1[1].Contains("n = 1")
                        && e2[1] == "    n = 1",
                        "加后=[" + e1[1] + "] 还原后=[" + e2[1] + "] want=[    n = 1]");

                    // 8. 用例 F：**多行选区**批量注释 —— "Ctrl+/ 失效"的真因：
                    //    Snapshot 对多行选区一律返回 null，于是每次按都判成"拿不到快照"，
                    //    键被吞、文档一个字不动（单选一直是好的）。选中两行 -> 两行都加撇号；
                    //    再按一次（选区会原样摆回去）-> 两行都还原。
                    RebuildScaffold(app, be, at, "    n = 1", "    n = 2");
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.SetSelectionRange(at + 1, 1, at + 2, 1 + "    n = 2".Length + 100);
                    });
                    Thread.Sleep(300);
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(800);
                    string[] f1 = ReadLines(app, be, at, 4);
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(800);
                    string[] f2 = ReadLines(app, be, at, 4);
                    check("F 多行选区 Ctrl+/ 批量注释（用户报的那一条）",
                        f1[1].TrimStart().StartsWith("'", StringComparison.Ordinal)
                        && f1[2].TrimStart().StartsWith("'", StringComparison.Ordinal)
                        && f2[1] == "    n = 1" && f2[2] == "    n = 2",
                        "加后=[" + f1[1] + "][" + f1[2] + "] 还原后=[" + f2[1] + "][" + f2[2] + "]");

                    // 9. 用例 G：**40 行**批量注释 —— 验 v145 换上的"VBE 原生执行"
                    //    在大选区上既不慢也不错。
                    //    旧实现：40 次 Lines(i,1) + 40 次 ReplaceLine + 1 次 SetSelection
                    //    = 81 次跨进程往返（真机 340ms 上下，用户就是这么感觉到"慢"的）；
                    //    现在：1 次 Lines(from,n) + 1 次 Execute + 0 次还原（选区原生保留）。
                    var many = new string[39];
                    for (int i = 1; i <= 39; i++) many[i - 1] = "    v" + i + " = " + i;
                    RebuildScaffold(app, be, at, "    v0 = 0", many);
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.SetSelectionRange(at + 1, 1, at + 40, 200);
                    });
                    Thread.Sleep(300);
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(900);
                    string[] g1 = ReadLines(app, be, at, 42);
                    int quoted = 0;
                    for (int i = 1; i <= 40; i++)
                        if (g1[i].TrimStart().StartsWith("'", StringComparison.Ordinal)) quoted++;
                    SendCombo(app, Native.VK_OEM_2, ctrl: true);
                    Thread.Sleep(900);
                    string[] g2 = ReadLines(app, be, at, 42);
                    int plain = 0;
                    for (int i = 1; i <= 40; i++)
                        if (!g2[i].TrimStart().StartsWith("'", StringComparison.Ordinal) && g2[i].Contains("v"))
                            plain++;

                    // 同一批 40 行，直接量"命令本身"的耗时（不掺键盘延迟）。
                    long msSet = -1, msUnset = -1;
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.SetSelectionRange(at + 1, 1, at + 40, 200);
                        long t = Environment.TickCount64;
                        Commands.ToggleComment(be);
                        msSet = Environment.TickCount64 - t;
                        long t2 = Environment.TickCount64;
                        Commands.ToggleComment(be);
                        msUnset = Environment.TickCount64 - t2;
                    });

                    check("G 40 行批量注释（一次 Execute，不再逐行）",
                        quoted == 40 && plain == 40 && msSet >= 0 && msSet < 200
                        && g1[40].TrimStart().StartsWith("'", StringComparison.Ordinal),
                        "注释 " + quoted + "/40 · 还原 " + plain + "/40 · 耗时 "
                        + msSet + "ms/" + msUnset + "ms（旧实现约 340ms）");

                    // 10. 用例 H：空格收起候选窗（用户报的 "set + 空格不收"）。
                    //     钩子对空格什么都不做时全靠轮询的 changed 兜，而 changed 比的是 Rstrip 之后的
                    //     文本：`Set ` 与 `Set` 一字不差 ⇒ 走不到 Trigger，列表就一直挂着。
                    //     这里用 Ctrl+Space 手动唤出（不受输入法状态影响），再发裸空格：既要验列表
                    //     收了，也要验空格照常落进文档（只收窗、不吞键）。
                    RebuildScaffold(app, be, at, "    n = 1");
                    // ★只发一次 Ctrl+Space 睡 600ms 就断言，真机约一半次数"唤不出来"，
                    //   日志里只有"唤出=False"，绿/红都不可信。现在：先读回光标/选区/快照
                    //   （多行选区会让 Snapshot 返回 null、Trigger 直接 Hide），唤不出来就
                    //   重摆光标再试，并把诊断串写进断言消息。
                    string hdiag = "";
                    bool shownH = false;
                    for (int attempt = 0; attempt < 3 && !shownH; attempt++)
                    {
                        app.RunOnUi(() =>
                        {
                            be.ForceReloadModule();
                            bool caretOk = be.SetCaret(at + 1, 1);
                            be.TryGetSelection(out int sl, out int sc, out int el, out int ec);
                            var s = be.Snapshot();
                            hdiag = "try" + (attempt + 1) + " caret=" + caretOk
                                    + " sel=" + sl + ":" + sc + "-" + el + ":" + ec
                                    + " snap=" + (s == null ? "null" : "L" + s.Line + "C" + s.CaretCol);
                        });
                        Thread.Sleep(250);
                        SendCombo(app, Native.VK_SPACE, ctrl: true);
                        Thread.Sleep(500);
                        shownH = app.PopupVisible;
                    }
                    SendKeys(app, Native.VK_SPACE);
                    Thread.Sleep(500);
                    bool hiddenH = !app.PopupVisible;
                    string[] h = ReadLines(app, be, at, 3);
                    check("H 空格收起候选窗（用户报的 set + 空格不收）",
                        shownH && hiddenH && h[1] == "     n = 1",
                        "唤出=" + shownH + " 空格后收起=" + hiddenH + " 行=[" + h[1] + "] " + hdiag);

                    // 11. 用例 I：**Ctrl+↑ / Ctrl+↓（不带 Shift）**。
                    //     带 Shift 的那组语义完全不同（↑ = 上方最近的声明行、↓ = 下方第一个
                    //     空白行），所以这两下一直漏给 VBE 原生。
                    //     这条同时能证明"键有没有被路由进新分支"：判据不成立会把键还回去
                    //     （Ctrl 被按住 ⇒ 实际发裸 ↑/↓），光标只走一行，断言必然对不上。
                    RebuildScaffold(app, be, at, "    x = 1",
                        "    Dim a As Long", "    y = 2", "    Dim b As Long", "    z = 3");
                    Thread.Sleep(300);
                    SendCombo(app, Native.VK_UP, ctrl: true);
                    Thread.Sleep(600);
                    int upLine = 0, upCol = 0;
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.TryGetSelection(out int l1, out int c1, out _, out _);
                        upLine = l1; upCol = c1;
                    });
                    SendCombo(app, Native.VK_DOWN, ctrl: true);
                    Thread.Sleep(600);
                    int dnLine = 0, dnCol = 0;
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.TryGetSelection(out int l2, out int c2, out _, out _);
                        dnLine = l2; dnCol = c2;
                    });
                    const string lastDimLine = "    Dim b As Long";   // at+4
                    const string lastCodeLine = "    z = 3";          // at+5
                    check("I Ctrl+↑/↓ 跳过程内两个端点",
                        upLine == at + 4 && upCol == lastDimLine.Length + 1
                        && dnLine == at + 5 && dnCol == lastCodeLine.Length + 1,
                        "↑=" + upLine + "," + upCol + " want " + (at + 4) + "," + (lastDimLine.Length + 1)
                        + "  ↓=" + dnLine + "," + dnCol + " want " + (at + 5) + "," + (lastCodeLine.Length + 1));
                }
                catch (Exception e)
                {
                    lock (gate) sb.AppendLine("EXCEPTION " + e);
                    allOk = false;
                }
                finally
                {
                    // 把输入法还给用户原来的状态（探针把它切到英数档才能逐字符打字）
                    Native.ImeRestore(Native.FocusedVbeWindow(), imeModeAtStart);
                    app.RequestExit();
                }
            })
            { IsBackground = true, Name = "enterprobe" };

            driver.Start();
            // ★整段清理放进 finally：探针中途抛异常（COM 抖动、脚手架建失败）时
            //   也绝不能把 `Sub VbcEnter()` 留在用户模块里。以前只有正常路径清理，
            //   一旦半路炸掉，垃圾就永久留在用户的代码里。
            bool joined = false;
            try
            {
                Application.Run(app);
            }
            finally
            {
                try
                {
                    joined = driver.Join(4000);
                    // ★这里不能再用 app.RunOnUi：Application.Run 已经返回，_disp 的
                    //   窗口没了，Invoke 会抛（被吞掉）—— 上一次就是因此把脚手架
                    //   留在了用户模块里。
                    be.ForceReloadModule();
                    int removed = CleanupScratch(be);
                    be.ForceReloadModule();
                    string now = be.ModuleText();
                    bool restored = string.Equals(now, orig, StringComparison.Ordinal);
                    if (!restored) allOk = false;
                    lock (gate)
                    {
                        sb.AppendLine("cleanup 块 = " + removed + "   restored = " + restored +
                                      "   lines " + be.LineCount() + "/" + baseCount +
                                      (joined ? "" : "   ⚠️ driver 线程超时未结束"));
                        sb.AppendLine(allOk ? "RESULT: PASS" : "RESULT: FAIL");
                        File.WriteAllText(Path.Combine(Settings.Home, "enterprobe.log"),
                            sb.ToString(), new UTF8Encoding(false));
                    }
                }
                catch (Exception e)
                {
                    // 清理本身失败：至少把"没清干净"喊出来，别让人以为模块是干净的
                    lock (gate)
                    {
                        sb.AppendLine("cleanup 异常：" + e.Message);
                        sb.AppendLine("RESULT: FAIL");
                        try
                        {
                            File.WriteAllText(Path.Combine(Settings.Home, "enterprobe.log"),
                                sb.ToString(), new UTF8Encoding(false));
                        }
                        catch { }
                    }
                    allOk = false;
                }
            }
            return allOk ? 0 : 1;
        }
    }
}
