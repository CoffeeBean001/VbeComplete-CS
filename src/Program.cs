// 入口：单实例 + 启动；带上 --selftest 时只跑自检、不装钩子（这样能在没有
// 桌面的环境里验证全部纯逻辑）。

using System;
using System.Collections.Generic;
using System.Drawing;
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
            // 诊断命令把中文打到控制台时，默认代码页（936）会吐一串乱码 ——
            // 这些输出同时会落盘成 UTF-8 日志，能对上才不会看错。
            try { Console.OutputEncoding = Encoding.UTF8; } catch { }

            if (args.Length > 0 && args[0] == "--selftest")
            {
                return SelfTest.Run();
            }
            if (args.Length > 0 && args[0] == "--probe")
            {
                return Probe();
            }
            // 只问"VBE 自带列表现在可见吗" —— 一次 ~2ms，不碰 COM，方便连测采样
            if (args.Length > 0 && args[0] == "--yieldprobe")
            {
                return YieldProbe();
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
            if (args.Length > 0 && args[0] == "--enterprobe")
            {
                return EnterProbe();
            }
            if (args.Length > 0 && args[0] == "--snipetest")
            {
                return SnippetTest();
            }
            if (args.Length > 0 && args[0] == "--dbtest")
            {
                return DbTest();
            }
            if (args.Length > 0 && args[0] == "--dump")
            {
                return DumpModule();
            }
            if (args.Length > 0 && args[0] == "--paretest")
            {
                return ParenTest();
            }
            if (args.Length > 0 && args[0] == "--enginetest")
            {
                return EngineTest();
            }
            if (args.Length > 0 && args[0] == "--startuptest")
            {
                return StartupTest();
            }
            // ★认不出的 `--xxx` 不能顺手当成"正常启动"：那会静默起一个托盘实例，
            //   让人以为"检查跑过了"（其实什么都没跑）。打一行用法就退出。
            if (args.Length > 0 && args[0].StartsWith("--", StringComparison.Ordinal))
            {
                const string usage =
                    "用法：VbeComplete.exe [开关]\r\n" +
                    "  （不带参数 = 正常启动）\r\n" +
                    "  --selftest    纯逻辑自检（不需要 Excel）\r\n" +
                    "  --dbtest      两个词库的建库/载入检查\r\n" +
                    "  --snipetest   片段库读写检查\r\n" +
                    "  --uitest      候选窗 / 片段窗离屏渲染出 PNG\r\n" +
                    "  --probe       打印当前 VBE 状态（只读）\r\n" +
                    "  --dump        把当前代码窗全文写到 moddump.txt（只读）\r\n" +
                    "  --enginetest  用真机模块文本跑候选判定（只读）\r\n" +
                    "  --paretest    真机验证函数补括号（改完即还原）\r\n" +
                    "  --carettest   真机验证中文名光标落点（改完即还原）\r\n" +
                    "  --writetest   真机验证写入链路（改完即还原）\r\n" +
                    "  --hktest      真机验证行操作快捷键（改完即还原）\r\n" +
                    "  --enterprobe  真机验证「回车自动补结构」（走真键盘钩子，改完即还原）\r\n" +
                    "  --startuptest 图标 + 开机自启的登记/回滚（会短暂改注册表，随后还原）\r\n";
                Console.WriteLine("未知开关：" + args[0]);
                Console.WriteLine(usage);
                try
                {
                    File.WriteAllText(Path.Combine(Settings.Home, "usage.log"), usage, Encoding.UTF8);
                }
                catch { }
                return 2;
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
        /// <summary>
        /// 把当前代码窗模块的全文写到 moddump.txt（只读，一个字节都不改）。
        /// 真机跑 --writetest / --hktest / --carettest 之前先 dump 一份，
        /// 跑完再 dump 一份 diff —— 那几条用例都会"改完再删"，这是唯一的核对手段。
        /// </summary>
        private static int DumpModule()
        {
            var be = new VbeBackend();
            string text = be.ModuleText();
            var snap = be.Snapshot();
            var sb = new StringBuilder();
            sb.AppendLine("module = " + (snap == null ? "?" : snap.Module) +
                          "  lines = " + be.LineCount() +
                          "  caret = " + (snap == null ? "?" : snap.Line + ":" + snap.CaretCol));
            sb.AppendLine("----");
            sb.Append(text ?? "(null)");
            File.WriteAllText(Path.Combine(Settings.Home, "moddump.txt"), sb.ToString(), Encoding.UTF8);
            return 0;
        }

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
        /// 括号补全验证（真机）：走 Engine.AcceptName 这条**真路径**（不是直接调
        /// Backend.ApplyCompletion），确认"选词库里的函数"真的会得到一对括号、
        /// 光标真的落在两个括号中间；类型名 / 关键字一个括号都不许补。
        ///
        /// ★为什么要单开一条：用户报"选中的是 vba 函数提示词，没有自动补全双括号"。
        ///   括号逻辑分散在四处（Followup 判档位 → IsCallSite 判位置 →
        ///   InCallStatement 判 Call 语句 → ApplyCompletion 落光标），只有真机
        ///   走一遍才知道断在哪一环。全程只动模块末尾的临时行，最后逐字节还原。
        /// </summary>
        /// <summary>日志用：安全截一段字符串（越界不抛，换行写成 \n）。</summary>
        private static string SafeSlice(string s, int from, int len)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (from < 0) from = 0;
            if (from >= s.Length) return "";
            string part = s.Substring(from, Math.Min(len, s.Length - from));
            return part.Replace("\r", "\\r").Replace("\n", "\\n");
        }

        /// <summary>
        /// 图标 + 开机自启的自检（不需要 Excel、不需要桌面）。
        ///   · 图标：内嵌资源在不在、九个尺寸能不能各自取到、拿到的像素尺寸对不对；
        ///   · 自启：登记 → 读回来要等于当前命令行 → 取消 → 读回来要没有。
        /// ★先给注册表和 autostart.txt 拍快照，测完原样还原 —— 不留任何痕迹。
        /// </summary>
        private static int StartupTest()
        {
            var sb = new StringBuilder();
            bool ok = true;
            Action<string, bool, string> check = (name, pass, detail) =>
            {
                if (!pass) ok = false;
                sb.AppendLine((pass ? "  PASS  " : "  FAIL  ") + name +
                              (string.IsNullOrEmpty(detail) ? "" : "   " + detail));
            };

            sb.AppendLine("=== icon + autostart test ===");

            // ---- 1. 图标 ----
            var asm = typeof(IconAssets).Assembly;
            bool hasRes = false;
            foreach (var n in asm.GetManifestResourceNames())
                if (n.EndsWith("app.ico", StringComparison.OrdinalIgnoreCase)) hasRes = true;
            check("内嵌 app.ico 资源存在", hasRes,
                "resources=" + string.Join(",", asm.GetManifestResourceNames()));

            int[] sizes = { 16, 20, 24, 32, 40, 48, 64, 128, 256 };
            var bad = new StringBuilder();
            foreach (int s in sizes)
            {
                try
                {
                    using (var ic = IconAssets.Get(s))
                    {
                        // Icon.Width 报的是"选中那一帧"的尺寸，可能不等于请求值。
                        // ★唯一已知的例外：256 那一帧是 PNG 压缩的（ICO 规范允许），
                        //   .NET 的 Icon(stream, size) 挑不到它，会退回 128 —— 我们
                        //   实际只用 16/20/32/48，256 是给资源管理器看 exe 用的，
                        //   所以这里只要求"别比请求的小太多"。
                        int want = Math.Min(s, 128);
                        if (ic == null || ic.Width < want || ic.Width > s)
                            bad.Append(s + "->" + (ic == null ? "null" : ic.Width.ToString()) + " ");
                    }
                }
                catch (Exception e) { bad.Append(s + "->异常:" + e.Message + " "); }
            }
            check("九个尺寸都能取到图标", bad.Length == 0, bad.ToString());

            // 托盘那一份必须不是"通用空白图标"（和 SystemIcons.Application 尺寸相同但内容不同，
            // 这里只做存在性判断；真正的观感看 src/icon-preview.png）
            using (var ti = IconAssets.Tray())
                check("托盘图标可用", ti != null && ti.Width >= 16, "w=" + (ti == null ? -1 : ti.Width));
            check("exe 路径可取", !string.IsNullOrEmpty(AutoStart.ExePath()), AutoStart.ExePath());

            // 把真正会挂到托盘上的那两个尺寸存成图片，肉眼确认"是不是那张蓝色 V"
            // （只看 log 里"加载成功"是不够的 —— 加载成功也可能是退回系统图标）
            try
            {
                using (var i32 = IconAssets.Get(32))
                using (var b32 = i32.ToBitmap())
                {
                    var canvas = new Bitmap(64 * 4 + 24, 80);
                    using (var g = Graphics.FromImage(canvas))
                    {
                        g.Clear(Color.FromArgb(32, 34, 38));
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
                        g.DrawImage(b32, new Rectangle(8, 8, 32, 32));
                        g.DrawImage(b32, new Rectangle(48, 8, 32, 32));
                        using (var i16 = IconAssets.Get(16))
                            g.DrawImage(i16.ToBitmap(), new Rectangle(88, 16, 16, 16));
                    }
                    canvas.Save(Path.Combine(Settings.Home, "icon-tray.png"),
                        System.Drawing.Imaging.ImageFormat.Png);
                    canvas.Dispose();
                }
            }
            catch (Exception e) { sb.AppendLine("  存 icon-tray.png 失败：" + e.Message); }

            // ---- 2. 开机自启 ----
            var snap = AutoStart.Capture();
            try
            {
                sb.AppendLine("  原始注册表值 = " + (snap.HadValue ? snap.Value : "（无）"));
                sb.AppendLine("  原始 " + "autostart.txt = " + (snap.HadMarker ? snap.Marker.Trim() : "（无）"));
                sb.AppendLine("  当前命令行 = " + AutoStart.CommandLine);

                bool on = AutoStart.SetEnabled(true);
                bool has = AutoStart.IsRegistered(out string v1);
                check("开启后注册表里有这一条", on && has, "value=" + (v1 ?? "null"));
                check("登记的内容就是本程序的命令行",
                    string.Equals(v1, AutoStart.CommandLine, StringComparison.OrdinalIgnoreCase),
                    v1 ?? "null");

                bool off = AutoStart.SetEnabled(false);
                bool has2 = AutoStart.IsRegistered(out string v2);
                check("关闭后注册表里没有了", off && !has2, "value=" + (v2 ?? "null"));
            }
            finally
            {
                AutoStart.Restore(snap);
            }
            bool restored = AutoStart.IsRegistered(out string v3);
            check("测完已还原成测试前的状态",
                restored == snap.HadValue &&
                (!snap.HadValue || string.Equals(v3, snap.Value, StringComparison.OrdinalIgnoreCase)),
                "value=" + (v3 ?? "null"));

            sb.AppendLine(ok ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "startuptest.log"), sb.ToString(), Encoding.UTF8);
            return ok ? 0 : 1;
        }

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

                // ==========================================================
                // 以下六条是 v132 新移植的"编辑手感"命令。每一条都自己造一小段
                // 临时代码、验完立刻整段删掉 —— 不依赖上面那 6 行的行号。
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
                    bool ok = r && Rules.Rstrip(now) == "    x = f" && snap != null && snap.CaretCol == 10;
                    check("退格删一对空括号", ok,
                        "ret=" + r + " 行=[" + now + "] caretCol=" + (snap?.CaretCol ?? -1) + " want=10");
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

                // ---- 13. Ctrl+Shift+L 重排缩进（有选区 -> 只理选区）----
                {
                    var blk = new[]
                    {
                        "Sub VbcFmt()", "Dim n As Long", "If n > 0 Then",
                        "n = n + 1", "Else", "n = n - 1", "End If", "End Sub",
                    };
                    int at = append(blk);
                    // ★必须拉一个真选区：没选区时我们理的是**整篇**，那就会动用户的代码
                    be.SetSelectionRange(at, 1, at + blk.Length - 1, blk[blk.Length - 1].Length + 1);
                    bool r = Commands.FormatIndent(be);
                    string f1 = be.LineAt(at + 1) ?? "";
                    string f2 = be.LineAt(at + 2) ?? "";
                    string f3 = be.LineAt(at + 3) ?? "";
                    string f4 = be.LineAt(at + 4) ?? "";
                    string f6 = be.LineAt(at + 6) ?? "";
                    bool ok = r
                              && f1 == "    Dim n As Long"
                              && f2 == "    If n > 0 Then"
                              && f3 == "        n = n + 1"
                              && f4 == "    Else"
                              && f6 == "    End If";
                    check("Ctrl+Shift+L 格式化缩进", ok,
                        "ret=" + r + " [" + f1 + "][" + f2 + "][" + f3 + "][" + f4 + "][" + f6 + "]");
                    dropFrom();
                }

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
        /// 真机"回车补结构"探针 —— **走真键盘钩子**。
        ///
        /// ★为什么非要有这一条：`--hktest` 是直接调 `VbeBackend.NewLineBelow`（等于
        ///   只验后端），而用户真正的路径是
        ///   `WH_KEYBOARD_LL → FreshHookSnap → EnterIndentWanted → Post → NewLineBelow`。
        ///   v137 就是栽在这上面：后端 15 条全 PASS，真机上回车照样不补 `Next`
        ///   （用户第二次报同一个问题）。所以这里把真 `App` 起起来（真钩子 + 真轮询），
        ///   用 SendInput 敲键，再回读模块——**验的是用户手指走的那条路**。
        ///
        /// 用例：
        ///   A) 光标停在现成的 `    For i = 1 To 3` 行尾，按一下回车；
        ///   B) 从空行**逐字符敲** `    For i = 1 To 3`，最后一个字符后 0ms 立刻回车
        ///      （模拟真人手速：这条专门盯着 `_pendingKeys` 那个"输入法组字"判据）；
        ///   C) 同上敲 `    If x > 0 Then`，验 `End If`。
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
                        a2[3] == "    Next" && a2[4] == "    x = 1",
                        "[" + a2[1] + "][" + a2[2] + "][" + a2[3] + "][" + a2[4] + "]  diag{" + diagA2 + "}");

                    // 4. 用例 B：空行上逐字符敲 For 行（模拟真人），最后一下立刻回车
                    RebuildScaffold(app, be, at, "");
                    Thread.Sleep(300);
                    bool typedB = TypeText(app, "    For i = 1 To 3", 45);
                    diagB = EnterDiag(app);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(700);
                    string[] b = ReadLines(app, be, at, 4);
                    check("B 逐字符敲 For + 立刻回车 -> 补 Next", typedB && b[3] == "    Next",
                        "typed=" + typedB + " " + Tail(b) + "  diag{" + diagB + "}");

                    // 5. 用例 C：If … Then -> End If
                    RebuildScaffold(app, be, at, "");
                    Thread.Sleep(300);
                    bool typedC = TypeText(app, "    If x > 0 Then", 45);
                    diagC = EnterDiag(app);
                    SendKeys(app, Native.VK_RETURN);
                    Thread.Sleep(700);
                    string[] c = ReadLines(app, be, at, 4);
                    check("C 逐字符敲 If … Then -> 补 End If", typedC && c[3] == "    End If",
                        "typed=" + typedC + " " + Tail(c) + "  diag{" + diagC + "}");

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
                        afterD == beforeD + 1 && d[1] == "    If x Then y = 1" && d[2].Trim() == "",
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
                }
                catch (Exception e)
                {
                    lock (gate) sb.AppendLine("EXCEPTION " + e);
                    allOk = false;
                }
                finally
                {
                    app.RequestExit();
                }
            })
            { IsBackground = true, Name = "enterprobe" };

            driver.Start();
            Application.Run(app);
            driver.Join(4000);

            // 清理 + 逐字节比对
            // ★这里不能再用 app.RunOnUi：Application.Run 已经返回，_disp 的窗口没了，
            //   Invoke 会抛（被吞掉）—— 上一次就是因此把脚手架留在了用户模块里。
            be.ForceReloadModule();
            int removed = CleanupScratch(be);
            be.ForceReloadModule();
            string now = be.ModuleText();
            bool restored = string.Equals(now, orig, StringComparison.Ordinal);
            if (!restored) allOk = false;
            lock (gate)
                sb.AppendLine("cleanup 块 = " + removed + "   restored = " + restored +
                              "   lines " + be.LineCount() + "/" + baseCount);

            lock (gate)
            {
                sb.AppendLine(allOk ? "RESULT: PASS" : "RESULT: FAIL");
                File.WriteAllText(Path.Combine(Settings.Home, "enterprobe.log"), sb.ToString(), new UTF8Encoding(false));
            }
            return allOk ? 0 : 1;
        }

        /// <summary>
        /// 删掉探针留下的 `Sub VbcEnter()` … `End Sub` 块（成功跑完要删、上一次
        /// 中途炸掉也要删 —— 探针绝不能把垃圾留在用户的模块里）。
        /// </summary>
        private static int CleanupScratch(VbeBackend be)
        {
            int removed = 0;
            for (int guard = 0; guard < 20; guard++)
            {
                be.ForceReloadModule();
                int cnt = be.LineCount();
                int head = -1, tail = -1;
                for (int i = cnt; i >= 1; i--)
                {
                    if (string.Equals((be.LineAt(i) ?? "").Trim(), "Sub VbcEnter()",
                        StringComparison.OrdinalIgnoreCase)) { head = i; break; }
                }
                if (head < 0) break;
                for (int i = head + 1; i <= Math.Min(cnt, head + 40); i++)
                {
                    if (string.Equals((be.LineAt(i) ?? "").Trim(), "End Sub",
                        StringComparison.OrdinalIgnoreCase)) { tail = i; break; }
                }
                if (tail < 0) break;
                be.DeleteLines(head, tail - head + 1);
                removed++;
            }
            return removed;
        }

        /// <summary>
        /// 把 at 之后的模块内容整段重建成：
        /// `Sub VbcEnter()` / <paramref name="header"/> / 逗号给的几行 / `End Sub`，
        /// 光标停在 header 行尾（header 为空则停在行首）。
        /// </summary>
        private static void RebuildScaffold(App app, VbeBackend be, int at, string header,
            params string[] extra)
        {
            app.RunOnUi(() =>
            {
                be.ForceReloadModule();
                int cnt = be.LineCount();
                if (cnt >= at) be.DeleteLines(at, cnt - at + 1);

                var body = new List<string> { "Sub VbcEnter()", header };
                body.AddRange(extra);
                body.Add("End Sub");
                for (int i = body.Count - 1; i >= 0; i--) be.InsertLines(at, body[i]);

                be.SetCaret(at + 1, header.Length + 1);
                be.InvalidateSnapshot();
            });
        }

        /// <summary>读回脚手架的前几行（探针用）。</summary>
        private static string[] ReadLines(App app, VbeBackend be, int at, int n)
        {
            string[] res = null;
            app.RunOnUi(() =>
            {
                be.ForceReloadModule();
                res = new string[n];
                for (int i = 0; i < n; i++) res[i] = be.LineAt(at + i) ?? "";
            });
            return res ?? new string[n];
        }

        private static string Tail(string[] lines)
        {
            var sb = new StringBuilder();
            for (int i = 1; i < lines.Length; i++) sb.Append("[").Append(lines[i]).Append("]");
            return sb.ToString();
        }

        /// <summary>钩子那一侧此刻看到的判据 —— 失败时把这一串打进日志，一眼看出卡在哪。</summary>
        private static string EnterDiag(App app)
        {
            var s = app.HookSnap;
            if (s == null) return "snap=null";
            string text = s.Text ?? "";
            bool wanted = Editing.EnterIndentWanted(text, s.CaretCol);
            return string.Format("popup={0} pending={1} age={2}ms L{3} col={4} dispEnd={5} want={6} line=[{7}]",
                app.PopupVisible, app.PendingKeys, app.HookSnapAgeMs, s.Line, s.CaretCol,
                Rules.DispCol(text, text.Length), wanted, text);
        }

        /// <summary>发一个键（只放行这一下注入）。</summary>
        private static void SendKeys(App app, int vk, bool shift = false)
        {
            App.AcceptInjectedVk = vk;
            try { Native.SendVk(vk, false, shift); }
            finally
            {
                Thread.Sleep(120);
                App.AcceptInjectedVk = 0;
            }
        }

        /// <summary>
        /// 发一个"组合键"（Ctrl+某键），只放行目标键这一下注入。
        /// ★Ctrl 的 down/up 也是注入的，会被钩子的 LLKHF_INJECTED 那道闸滤掉 ——
        ///   但它们照样会进系统，所以 Down(VK_CONTROL) 在目标键到达时是真的。
        ///   目标键本身通过 AcceptInjectedVk 单点放行，钩子走的就是用户手按那条路。
        /// </summary>
        private static void SendCombo(App app, int vk, bool ctrl = false, bool shift = false)
        {
            App.AcceptInjectedVk = vk;
            try { Native.SendVk(vk, ctrl, shift); }
            finally
            {
                Thread.Sleep(150);
                App.AcceptInjectedVk = 0;
            }
        }

        /// <summary>逐字符敲一串文本（用真 vk，模拟真人按键；注入标记在探针里放行）。</summary>
        private static bool TypeText(App app, string s, int delayMs)
        {
            App.AcceptInjectedVk = -1;
            try
            {
                foreach (char ch in s)
                {
                    if (!CharToVk(ch, out int vk, out bool shift)) return false;
                    Native.SendVk(vk, false, shift);
                    Thread.Sleep(delayMs);
                }
                return true;
            }
            finally
            {
                App.AcceptInjectedVk = 0;
            }
        }

        private static bool CharToVk(char c, out int vk, out bool shift)
        {
            shift = false;
            if (c >= 'a' && c <= 'z') { vk = c - 'a' + 'A'; return true; }
            if (c >= 'A' && c <= 'Z') { vk = c; shift = true; return true; }
            if (c >= '0' && c <= '9') { vk = c; return true; }
            switch (c)
            {
                case ' ': vk = 0x20; return true;
                case '=': vk = 0xBB; return true;                 // VK_OEM_PLUS
                case '+': vk = 0xBB; shift = true; return true;
                case '-': vk = 0xBD; return true;                 // VK_OEM_MINUS
                case '.': vk = 0xBE; return true;                 // VK_OEM_PERIOD
                case '>': vk = 0xBE; shift = true; return true;
                case ',': vk = 0xBC; return true;                 // VK_OEM_COMMA
                case '<': vk = 0xBC; shift = true; return true;
                case '"': vk = 0xDE; shift = true; return true;   // VK_OEM_7
                case '\'': vk = 0xDE; return true;
            }
            vk = 0;
            return false;
        }

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
        /// 候选判定的**真机**回归：拿真实模块文本 + 人造快照跑 Trigger，
        /// 逐条验证"这个位置到底弹不弹、弹出的第一条是什么"。
        ///
        /// ★为什么不直接敲键盘：钩子按设计不理注入键，没法在这里模拟打字。
        ///   但 Trigger 的全部输入就是 (模块文本, 行, 列)，人造快照与真打字
        ///   走的是同一条代码路径 —— 判定逻辑 100% 覆盖，只有"按键有没有送到
        ///   钩子"覆盖不到（那部分由 --hktest 的命令级用例兜）。
        ///
        /// 本命令纯读，一个字符都不写模块。
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
                "    Dim 他们 As String\n" +
                "    Dim dic As Object, c As cls1\n" +
                "    student = getOne()\n" +
                "    \n" +
                "End Sub\n";
            const string modCls =
                "Option Explicit\n" +
                "\n" +
                "Public s1 As String\n" +
                "Public s2 As String\n";

            var synth = new SyntheticBackend()
                .Add("模块1", 1, modStd)
                .Add("cls1", 2, modCls);

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

            // ---- 2. Option Explicit：未声明的名字不进候选 ----
            // 光标所在行会被"挖掉半截词"，所以搜索词要和模块里的名字错开：
            // 池子里若混进了第 10 行的 `student`，它就会在 stud 这次冒出来。
            sprobe("未声明 student 不进候选（stud|）", "模块1", 11,
                "    stud", 9, null, new[] { "student2" }, new[] { "student" });
            sprobe("未声明 ok 不进候选（ok|）", "模块1", 11,
                "    ok", 7, null, new[] { "okkk2" }, new[] { "ok" });
            sprobe("中文变量名能提示（他|）", "模块1", 11,
                "    他", 7, true, new[] { "他们" }, null);

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
        /// 用来钉"跨模块可见性"这类**与当前打开哪个模块无关**的契约 ——
        /// 真机探针在没开代码窗时一律 skip，等于什么都没验（v137 之前
        /// paretest / hktest / enginetest 三条就是这么"绿"着的）。
        /// </summary>
        private sealed class SyntheticBackend : IVbeSource
        {
            private readonly Dictionary<string, string> _texts =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, int> _types =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            public string Current;

            public SyntheticBackend Add(string name, int type, string text)
            {
                _texts[name] = text;
                _types[name] = type;
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
                    list.Add(ms);
                }
                return list;
            }

            public bool ApplyCompletion(int lineNo, int wordStartCol, int endCol,
                string completion, bool wantParens, bool wantSpace, string enclosingProc) => false;

            public string EnclosingProcName() => "test";
        }

        private static string LineOf(string text, int line1)
        {
            if (text == null || line1 < 1) return "";
            var all = Parser.SplitLines(text);
            return line1 <= all.Length ? all[line1 - 1] : "";
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
                // 用**真实匹配 + 真实排序**造数据：预览图里看到的高亮位置、
                // 图标档位、行序，都与用户在 VBE 里看到的一致。
                // 手写 Pos/Span 的那种预览图，改了算法它也不会跟着变。
                // ★词库也要真的载入：否则预览图里的内建名字全变成"var"图标。
                var wl = new WordLibrary();
                wl.Load();
                sb.AppendLine("word db: kw=" + wl.KeywordCount + " fn=" + wl.FunctionCount +
                              " kwPath=" + WordLibrary.KeywordPath() +
                              "(exists=" + File.Exists(WordLibrary.KeywordPath()) + ")" +
                              " err=" + wl.LastError);
                var eng = new Engine { Recent = new Recent() };
                eng.PrimePool("", "Module1");
                const string word = "se";
                var cands = new System.Collections.Generic.List<Candidate>();
                foreach (var n in new[]
                         {
                             "Select", "selBegin", "Size", "GetSettingTitle3", "dataSheet",
                             "MsgBox", "String", "Set", "Screen", "ScrollBars",
                         })
                {
                    FuzzyHit fh;
                    if (!Fuzzy.Match(n, word, out fh)) continue;
                    cands.Add(new Candidate
                    {
                        Name = n,
                        Kind = eng.Classify(n),
                        Pos = fh.First,
                        Span = fh.Span,
                        Hits = fh.Positions,
                        HitKind = fh.Kind,
                        Run = fh.Run,
                        ScopeKind = 3,
                        IsBuiltin = VbaBuiltins.IsKeyword(n),
                    });
                }
                cands.Sort(Engine.Cmp);

                f.SetData(cands, 1, 0, 15);
                var size = f.DesiredSize();
                int w = size.Width;
                int h = size.Height;
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
                    sb.AppendLine("saved " + path + "  " + w + "x" + h +
                                  "  items=" + cands.Count + "  word=" + word +
                                  "  font=" + f.FontName + "  rowH=" + f.RowHeight);
                }
                f.Hide();
                f.Dispose();

                // ---- 片段窗也离屏渲染一张：布局/配色/按钮顺序一眼能核对 ----
                try
                {
                    using (var store = new SnippetStore())
                    {
                        var sf = new SnippetForm(store, code => false);
                        sf.StartPosition = FormStartPosition.Manual;
                        sf.Location = new System.Drawing.Point(-4000, -4000);
                        sf.Show();
                        sf.PreviewOnly(store.All());
                        Application.DoEvents();
                        int sw = sf.Width, sh = sf.Height;
                        using (var bmp2 = new System.Drawing.Bitmap(sw, sh))
                        {
                            sf.DrawToBitmap(bmp2, new System.Drawing.Rectangle(0, 0, sw, sh));
                            string p2 = Path.Combine(Settings.Home, "snippet-preview.png");
                            bmp2.Save(p2, System.Drawing.Imaging.ImageFormat.Png);
                            sb.AppendLine("saved " + p2 + "  " + sw + "x" + sh +
                                           "  rows=" + store.Count());
                        }
                        sf.Hide();
                        sf.Dispose();
                    }
                }
                catch (Exception e2)
                {
                    sb.AppendLine("snippet preview skipped: " + e2.Message);
                }

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
            // v139：让位判据（App.PollOnce 每轮都问这一个函数）
            try
            {
                sb.AppendLine("vbeMemberListVisible = " + Native.VbeMemberListVisible());
                var rects = Native.VbePopupRects();
                sb.AppendLine("vbePopupRects = " + (rects == null || rects.Length == 0 ? "(none)" :
                    string.Join(" | ", rects.Select(r => r.Left + "," + r.Top + "-" + r.Right + "," + r.Bottom))));
            }
            catch (Exception e) { sb.AppendLine("vbeMemberListVisible ERR " + e.Message); }
            File.WriteAllText(Path.Combine(Settings.Home, "probe-cs.log"), sb.ToString(), Encoding.UTF8);
            return 0;
        }

        /// <summary>
        /// 只采样一次"让位判据"（App.PollOnce 用的同一个函数）。
        /// 不碰 COM，一次约 2ms —— 真机上要连测"VBE 的列表出现时我们能不能看见"，
        /// 走 --probe 太重（要过一遍 COM，几秒钟）。
        /// </summary>
        private static int YieldProbe()
        {
            var sb = new StringBuilder();
            int visible = 0, seen = 0;
            for (int i = 0; i < 5; i++)
            {
                seen++;
                bool vis = Native.VbeMemberListVisible();
                if (vis) visible++;
                var rects = Native.VbePopupRects();
                sb.AppendLine(DateTime.Now.ToString("HH:mm:ss.fff") +
                              " memberListVisible=" + vis +
                              " rects=" + (rects == null || rects.Length == 0 ? "(none)" :
                                  string.Join(" | ", rects.Select(r => r.Left + "," + r.Top + "-" + r.Right + "," + r.Bottom))));
                Thread.Sleep(200);      // 缓存 150ms，睡久一点才能读到新采样
            }
            sb.AppendLine("RESULT: " + (visible > 0 ? "VBE-LIST-VISIBLE" : "VBE-LIST-NOT-VISIBLE") +
                          " (" + visible + "/" + seen + ")");
            File.WriteAllText(Path.Combine(Settings.Home, "yieldprobe.log"), sb.ToString(), Encoding.UTF8);
            return visible > 0 ? 0 : 1;
        }

        private static string App_invbe()        {
            IntPtr frame = Native.FindWindow("wndclass_desked_gsk", null);
            if (frame == IntPtr.Zero) return "no vbe frame";
            uint pid;
            uint tid = Native.GetWindowThreadProcessId(frame, out pid);
            IntPtr f = Native.FocusedHwndOf(tid);
            return "frame=" + frame + " pid=" + pid + " focusClass=" + Native.ClassNameOf(f);
        }
    }
}
