// 入口：单实例 + 启动；带上 --selftest 时只跑自检、不装钩子（这样能在没有
// 桌面的环境里验证全部纯逻辑）。

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace VbeComplete
{
    internal static class Program
    {
        private static Mutex _mutex;

        /// <summary>已经有一份在跑时，用它叫那一份把托盘图标提示出来。</summary>
        public const string ShowEventName = @"Local\VbeComplete_CSharp_v1_show";

        [STAThread]
        private static int Main(string[] args)
        {
            // 崩溃必须留下痕迹。WinExe 没有控制台，CLR 把异常吞掉之后用户只看到
            // "窗口闪退、项目没启动"，我们手里连一行线索都没有。
            AppDomain.CurrentDomain.UnhandledException +=
                (s, e) => Log.Crash("AppDomain.UnhandledException", e.ExceptionObject as Exception);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException +=
                (s, e) => Log.Crash("Application.ThreadException", e.Exception);

            try { return RunMain(args); }
            catch (Exception e)
            {
                Log.Crash("Main", e);
                try
                {
                    MessageBox.Show(
                        "启动失败：" + e.Message + "\r\n\r\n详情见程序目录下的 crash.log",
                        "VbeComplete", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
                return 1;
            }
        }

        private static int RunMain(string[] args)
        {
            Settings.Load();

            if (args.Length > 0 && args[0] == "--selftest")
            {
                return SelfTest.Run();
            }
            if (args.Length > 0 && args[0] == "--probe")
            {
                return Probe();
            }
            if (args.Length > 0 && args[0] == "--writetest")
            {
                return WriteTest();
            }
            if (args.Length > 0 && args[0] == "--uitest")
            {
                return UiTest();
            }
            if (args.Length > 0 && args[0] == "--carettest")
            {
                return CaretTest();
            }
            if (args.Length > 0 && args[0] == "--hktest")
            {
                return HkTest();
            }

            // 单实例：同名互斥体已经存在就说明有一份在跑。
            bool created;
            _mutex = new Mutex(true, "VbeComplete_CSharp_v1", out created);
            if (!created)
            {
                // 只写日志是不够的 —— 用户看到的就是"双击了，什么都没发生"。
                // 叫已经在跑的那一份弹个气泡，把托盘图标指出来。
                SignalExisting();
                Log.Write("已有一份在运行，本次退出（已提示已有实例）");
                return 0;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var app = new App();
            app.Start();

            // Ctrl+C / 关控制台 时收干净
            Console.CancelKeyPress += (s, e) => { e.Cancel = true; app.Shutdown(); };

            Application.Run(app);
            app.Shutdown();
            try { _mutex.ReleaseMutex(); } catch { }
            return 0;
        }

        private static void SignalExisting()
        {
            EventWaitHandle h;
            if (EventWaitHandle.TryOpenExisting(ShowEventName, out h))
            {
                try { h.Set(); } catch { }
                h.Dispose();
            }
        }

        /// <summary>
        /// 写入链路验证：在模块【末尾】追加一行测试代码 -> 用补全改它 -> 删掉。
        /// 全程可逆，最后逐字节比对原文；比对失败会把原文整段写进日志。
        /// ★只由用户显式运行 `--writetest` 触发，正常运行一次都不会动用户代码。
        /// </summary>
        private static int WriteTest()
        {
            var sb = new StringBuilder();
            var be = new VbeBackend();
            string orig = be.ModuleText();
            int count = be.LineCount();
            bool ok = true;

            sb.AppendLine("=== write test ===");
            sb.AppendLine("module lines = " + count + "  text len = " + (orig == null ? -1 : orig.Length));
            if (count <= 0)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "writetest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }

            try
            {
                const string probe = "    Dim abc";
                if (!be.InsertLines(count + 1, probe))
                {
                    sb.AppendLine("InsertLines FAILED");
                    ok = false;
                }
                else
                {
                    string back = be.LineAt(count + 1);
                    sb.AppendLine("insert readback = [" + back + "]");
                    bool same = back == probe;
                    if (!same) ok = false;
                    sb.AppendLine("insert roundtrip = " + same);

                    bool w = be.ApplyCompletion(count + 1, 9, 12, "abcd", false, false, null);
                    string after = be.LineAt(count + 1);
                    sb.AppendLine("apply ok=" + w + "  result=[" + after + "]");
                    sb.AppendLine("apply result correct = " + (after == "    Dim abcd"));
                    if (after != "    Dim abcd") ok = false;

                    be.SetCaret(count + 1, 13);
                    var s2 = be.Snapshot();
                    string got = s2 == null ? "null" : s2.Line + ":" + s2.CaretCol;
                    sb.AppendLine("caret after SetCaret(13) = " + got + " (want 13)");
                    if (s2 == null || s2.CaretCol != 13) ok = false;
                }
            }
            finally
            {
                // 无论成败都要把测试行删干净（最多重试 3 次）
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
                sb.AppendLine("!!! 未能逐字节还原，原文如下（请手工检查模块末尾） !!!");
                sb.AppendLine(orig);
                sb.AppendLine("!!! 现在的全文 !!!");
                sb.AppendLine(now);
            }
            sb.AppendLine(ok && restored ? "RESULT: PASS" : "RESULT: FAIL");

            File.WriteAllText(Path.Combine(Settings.Home, "writetest.log"), sb.ToString(), Encoding.UTF8);
            return ok && restored ? 0 : 1;
        }

        /// <summary>
        /// 光标落点验证（真机）：模拟"打到一半 + Tab 补全"，读回 VBE 真实光标列，
        /// 跟"按读回来的真行算出来的列"对数，并把光标左右两截文本打出来。
        /// ★只在 `--carettest` 时执行；动的是模块末尾的临时行，finally 里删干净
        /// 并逐字节比对原文。
        /// </summary>
        /// <summary>
        /// 真机验证五个命令：Ctrl+Shift+Enter / Ctrl+D / Ctrl+K / Ctrl+/ / Ctrl+Shift+↑↓。
        ///
        /// ★为什么要这个：用户报"快捷键全都没了"。按键有没有走到钩子，在沙箱里
        ///   没法用真键盘试（钩子按设计不理注入键）；但"命令本身能不能干活"可以
        ///   100% 验证 —— 因为这里调的就是钩子调的那几个 Commands.* 函数。
        ///   命令全过 ⇒ 问题只可能在路由（焦点判定 / 修饰键 / 别的软件抢钩子），
        ///   App 里那两行"热键被跳过 / 热键未命中"日志就是给这种情况准备的。
        ///
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

            int Lhead = baseCount + 1, Ldim = baseCount + 2, L1 = baseCount + 3,
                L2 = baseCount + 4, Lblank = baseCount + 5, Lend = baseCount + 6;

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

                // ---- 5. Ctrl+Shift+↑ 落到最后一个声明行 ----
                {
                    be.SetCaret(L2, 1);
                    be.InvalidateSnapshot();
                    bool r = Commands.CtrlNavJump(be, -1);
                    var snap = be.Snapshot();
                    bool ok = r && snap != null && snap.Line == Ldim;
                    check("Ctrl+Shift+↑ 落最后 Dim", ok,
                        "ret=" + r + " caretLine=" + (snap?.Line ?? -1) + " want=" + Ldim);
                }

                // ---- 6. Ctrl+Shift+↓ 跳到收尾空白行 ----
                {
                    be.SetCaret(L1, 1);
                    be.InvalidateSnapshot();
                    bool r = Commands.CtrlNavJump(be, +1);
                    var snap = be.Snapshot();
                    bool ok = r && snap != null && snap.Line == Lblank;
                    // 解析结果打出来：行号对不上时，一眼就能看出是"解析偏了一行"
                    string mt = be.ModuleText() ?? "";
                    var pl = Parser.SplitLines(mt);
                    var pp = Parser.ExtractProcedures(mt);
                    sb.AppendLine("        parsed lines=" + pl.Length + " procs=" + pp.Count +
                                  " blank@[" + string.Join(",",
                                      Enumerable.Range(1, pl.Length)
                                          .Where(i => Parser.CodeOnly(pl[i - 1]).Trim().Length == 0)
                                          .Select(i => i.ToString())) + "]");
                    check("Ctrl+Shift+↓ 跳空白行", ok,
                        "ret=" + r + " caretLine=" + (snap?.Line ?? -1) + " want=" + Lblank);
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

        private static int CaretTest()
        {
            var sb = new StringBuilder();
            var be = new VbeBackend();
            string orig = be.ModuleText();
            int count = be.LineCount();
            bool allOk = true;

            sb.AppendLine("=== caret test ===");
            sb.AppendLine("module lines = " + count + "  text len = " + (orig == null ? -1 : orig.Length));
            if (count <= 0)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "carettest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }

            // (标签, 插入的行, 已输入的词, 候选)
            // ⚠️ InsertLines 自己也会触发 VBE 的行格式化（`student=spl` 读回是
            //   `student = spl`），所以列**必须按读回来的真行**算 —— 这正是
            //   真机与离线用例的差别，也是这个 bug 之前没被发现的原因。
            var cases = new[]
            {
                ("顶格无缩进",        "student=spl",  "spl", "Split"),
                ("4 空格缩进",        "    student=spl", "spl", "Split"),
                ("Tab 缩进",          "\tstudent=spl", "spl", "Split"),
                ("Tab 缩进 + 全角名",  "\t学生=spl",   "spl", "Split"),
                ("右边还有同名子串",   "    abc=spl+Splitter", "spl", "Split"),
            };

            try
            {
                foreach (var c in cases)
                {
                    int line = count + 1;
                    if (!be.InsertLines(line, c.Item2))
                    {
                        sb.AppendLine(c.Item1 + ": InsertLines FAILED");
                        allOk = false;
                        continue;
                    }
                    string inserted = be.LineAt(line) ?? "";
                    int at0 = WholeWordIndex(inserted, c.Item3);
                    if (at0 < 0)
                    {
                        sb.AppendLine(c.Item1 + ": 读回行里找不到已输入的词 [" + inserted + "]");
                        allOk = false;
                        be.DeleteLines(line, 1);
                        continue;
                    }
                    int wordStart = at0 + 1;                  // 1-based 字符列
                    int endCol = at0 + 1 + c.Item3.Length;

                    // ★v131c：这一条验的是【Engine 真正用的那对换算】——
                    //   VBE 的 `CaretCol` 是显示列，直接拿去切字符下标，含全角/Tab 的行
                    //   会切出空词，症状是"这一行根本不弹候选"。所以用"打完还没补全"
                    //   的那一行做样本：把词尾的显示列换算成字符列再取词，必须还是原词。
                    int wordEndDisp = Rules.DispCol(inserted, at0 + c.Item3.Length, 4);
                    int ws;
                    string wordSeen = Rules.WordBefore(inserted,
                        Rules.CharColFromDisp(inserted, wordEndDisp), out ws);
                    bool okWord = string.Equals(wordSeen, c.Item3, StringComparison.OrdinalIgnoreCase);

                    bool w = be.ApplyCompletion(line, wordStart, endCol, c.Item4, true, false, null);
                    string actual = be.LineAt(line) ?? "";
                    var snap = be.Snapshot();
                    int got = snap == null ? -1 : snap.CaretCol;

                    // ★v131c：判据必须【独立】—— 上一版用 DispCol 算期望列，等于拿
                    //   被测的同一套语义去验它自己，于是"落在 `(` 左边"也能全绿。
                    //   现在直接看光标左右两边是什么字符：
                    //   补过括号 ⇒ 左邻必须是 `(`、右邻必须是 `)`（= 夹在括号中间）。
                    int idx = actual.IndexOf('(');
                    bool hasParens = idx >= 0;
                    int want = hasParens
                        ? Rules.DispCol(actual, idx + 1, 4)          // `)` 之前
                        : Rules.DispCol(actual, at0 + c.Item4.Length, 4);

                    // 约定（与 Python 版 `_set_caret` 一致）：列 C = "字符 C-1 之前"，
                    // ★但 C 是【显示列】—— 全角/ Tab 行上不能直接减 1，要先换回字符列。
                    int at = Rules.CharColFromDisp(actual, got) - 1;
                    string left = (at < 0 || at > actual.Length) ? "?" : actual.Substring(0, at);
                    string right = (at < 0 || at > actual.Length) ? "?" : actual.Substring(at);
                    char chLeft = (at >= 1 && at <= actual.Length) ? actual[at - 1] : '\0';
                    char chRight = (at >= 0 && at < actual.Length) ? actual[at] : '\0';
                    bool betweenParens = hasParens && chLeft == '(' && chRight == ')';

                    bool ok = w && got == want && okWord && betweenParens;
                    if (!ok) allOk = false;

                    sb.AppendLine("--- " + c.Item1 + " ---");
                    sb.AppendLine("  inserted  = [" + inserted.Replace("\t", "\\t") + "]");
                    sb.AppendLine("  actual    = [" + actual.Replace("\t", "\\t") + "]");
                    sb.AppendLine("  apply ok  = " + w + "   (wordStart=" + wordStart + " endCol=" + endCol + ")");
                    sb.AppendLine("  caret col = " + got + "   want = " + want + "   " + (got == want ? "OK" : "MISMATCH"));
                    sb.AppendLine("  parens    = 左[" + Esc(chLeft) + "] 右[" + Esc(chRight) + "]  " +
                                  (betweenParens ? "OK(夹在括号中间)" : (hasParens ? "MISMATCH(不在括号中间)" : "无括号")));
                    sb.AppendLine("  engine取词= [" + wordSeen + "]  " + (okWord ? "OK" : "MISMATCH"));
                    sb.AppendLine("  left      = [" + left.Replace("\t", "\\t") + "]");
                    sb.AppendLine("  right     = [" + right.Replace("\t", "\\t") + "]");

                    be.DeleteLines(line, 1);
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
            if (!restored) allOk = false;
            sb.AppendLine(allOk ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "carettest.log"), sb.ToString(), Encoding.UTF8);
            return allOk ? 0 : 1;
        }

        /// <summary>把单个字符渲染成日志里看得见的样子（空格/Tab/空位都要能区分）。</summary>
        private static string Esc(char c) =>
            c == '\0' ? "∅" : c == '\t' ? "\\t" : c == ' ' ? "␣" : c.ToString();

        /// <summary>整词查找（前后都不是标识符字符），找不到返回 -1。测试用。</summary>
        private static int WholeWordIndex(string text, string word)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(word)) return -1;
            int i = 0;
            while (i <= text.Length - word.Length)
            {
                int hit = text.IndexOf(word, i, StringComparison.OrdinalIgnoreCase);
                if (hit < 0) return -1;
                bool okL = hit == 0 || !Rules.IsIdentChar(text[hit - 1]);
                bool okR = hit + word.Length >= text.Length || !Rules.IsIdentChar(text[hit + word.Length]);
                if (okL && okR) return hit;
                i = hit + 1;
            }
            return -1;
        }

        /// <summary>
        /// 候选窗离屏渲染：把窗口画进一张 PNG，用来一眼核对配色 / 图标 / 布局。
        /// 不需要桌面，也不会出现在屏幕上。
        /// </summary>
        private static int UiTest()
        {
            var sb = new StringBuilder();
            try
            {
                var f = new PopupForm();
                var items = new System.Collections.Generic.List<Candidate>
                {
                    new Candidate { Name = "MsgBox",   Kind = "func", Pos = 0, Span = 2 },
                    new Candidate { Name = "Main",     Kind = "func", Pos = 0, Span = 2 },
                    new Candidate { Name = "mCount",   Kind = "var",  Pos = 1, Span = 1 },
                    new Candidate { Name = "Size",     Kind = "prop", Pos = 0, Span = 2 },
                    new Candidate { Name = "String",   Kind = "type", Pos = 0, Span = 2 },
                    new Candidate { Name = "Select",   Kind = "kw",   Pos = 0, Span = 2 },
                    new Candidate { Name = "selBegin", Kind = "var",  Pos = 0, Span = 2 },
                };
                f.SetData(items, 1, 0, 8);
                int w = 260;
                int h = items.Count * f.RowHeight + f.DetailHeight + 2;
                f.StartPosition = FormStartPosition.Manual;
                f.Location = new System.Drawing.Point(-4000, -4000);
                f.Show();
                f.Width = w; f.Height = h;
                Application.DoEvents();

                using (var bmp = new System.Drawing.Bitmap(w, h))
                {
                    f.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, w, h));
                    string path = Path.Combine(Settings.Home, "popup-preview.png");
                    bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                    sb.AppendLine("saved " + path + "  " + w + "x" + h);
                }
                f.Hide();
                f.Dispose();
                sb.AppendLine("RESULT: PASS");
                File.WriteAllText(Path.Combine(Settings.Home, "uitest.log"), sb.ToString(), Encoding.UTF8);
                return 0;
            }
            catch (Exception e)
            {
                sb.AppendLine("EXCEPTION " + e);
                File.WriteAllText(Path.Combine(Settings.Home, "uitest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }
        }

        /// <summary>诊断用：把"能不能连上 VBE"写进 probe.log。</summary>
        private static int Probe()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Home = " + Settings.Home);
            sb.AppendLine("settings.ini = " + Path.Combine(Settings.Home, "settings.ini"));
            sb.AppendLine("settings error = " + Settings.LastError());
            sb.AppendLine("kernel = " + Environment.Version);
            sb.AppendLine("64bit = " + Environment.Is64BitProcess);

            var be = new VbeBackend();
            sb.AppendLine("--- COM ---");
            foreach (var progid in new[] { "Excel.Application", "Word.Application" })
            {
                object app = ComCtx.GetActive(progid);
                sb.AppendLine(progid + " active = " + (app == null ? "null" : "ok"));
                if (app == null) continue;
                try
                {
                    dynamic dapp = app;
                    dynamic vbe = dapp.VBE;
                    sb.AppendLine("  VBE = " + (vbe == null ? "null" : "ok"));
                    if (vbe == null) continue;
                    try
                    {
                        dynamic projs = vbe.VBProjects;
                        int pc = Convert.ToInt32(projs.Count);
                        sb.AppendLine("  projects = " + pc);
                        if (pc > 0)
                        {
                            dynamic p0 = projs.Item(1);
                            sb.AppendLine("  project[1].Name = " + Convert.ToString(p0.Name));
                            dynamic comps = p0.VBComponents;
                            sb.AppendLine("  components = " + Convert.ToInt32(comps.Count));
                        }
                    }
                    catch (Exception e) { sb.AppendLine("  projects ERR " + e.Message); }
                    try
                    {
                        dynamic cp = vbe.ActiveCodePane;
                        sb.AppendLine("  ActiveCodePane = " + (cp == null ? "null" : "ok"));
                        if (cp != null)
                        {
                            dynamic cm = cp.CodeModule;
                            sb.AppendLine("  module = " + Convert.ToString(cm.Name) +
                                          " lines = " + Convert.ToInt32(cm.CountOfLines));
                            try
                            {
                                dynamic sel = cp.Selection;
                                sb.AppendLine("  Selection = " +
                                    Convert.ToInt32(sel.StartLine) + "," + Convert.ToInt32(sel.StartColumn) +
                                    " .. " + Convert.ToInt32(sel.EndLine) + "," + Convert.ToInt32(sel.EndColumn));
                            }
                            catch (Exception e) { sb.AppendLine("  Selection ERR " + e.Message); }
                            try
                            {
                                object g = cp.GetSelection();
                                sb.AppendLine("  GetSelection = " + (g == null ? "null" : g.GetType().FullName));
                            }
                            catch (Exception e) { sb.AppendLine("  GetSelection ERR " + e.Message); }
                        }
                    }
                    catch (Exception e) { sb.AppendLine("  ActiveCodePane ERR " + e.Message); }
                }
                catch (Exception e) { sb.AppendLine("  VBE ERR " + e.Message); }
            }

            sb.AppendLine("--- backend ---");
            var snap = be.Snapshot();
            sb.AppendLine("snapshot = " + (snap == null ? "null" : snap.ToString()));
            if (snap != null)
            {
                sb.AppendLine("line text = " + (snap.Text ?? ""));
                sb.AppendLine("module text len = " + (be.ModuleText()?.Length ?? -1));
                sb.AppendLine("components = " + string.Join(", ", be.ComponentNames()));
            }
            sb.AppendLine("--- settings ---");
            sb.AppendLine("AUTO_SPACE = " + Settings.Flag("NO_AUTOSPACE", false));
            sb.AppendLine("HIDE_ON_COMPLETE = " + Settings.Flag("HIDE_ON_COMPLETE", true));
            sb.AppendLine("VIEW_ROWS = " + Settings.Int("VIEW_ROWS", 8));
            sb.AppendLine("--- win32 ---");
            sb.AppendLine("inVbeCodePane(win32) = " + App_invbe());
            File.WriteAllText(Path.Combine(Settings.Home, "probe-cs.log"), sb.ToString(), Encoding.UTF8);
            return 0;
        }

        private static string App_invbe()
        {
            IntPtr frame = Native.FindWindow("wndclass_desked_gsk", null);
            if (frame == IntPtr.Zero) return "no vbe frame";
            uint pid;
            uint tid = Native.GetWindowThreadProcessId(frame, out pid);
            IntPtr f = Native.FocusedHwndOf(tid);
            return "frame=" + frame + " pid=" + pid + " focusClass=" + Native.ClassNameOf(f);
        }
    }
}
