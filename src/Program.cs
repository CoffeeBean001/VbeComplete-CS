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
    internal static partial class Program
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
            bool openSettings = false;
            Settings.Load();
            // 诊断命令把中文打到控制台时，默认代码页（936）会吐一串乱码 ——
            // 这些输出同时会落盘成 UTF-8 日志，能对上才不会看错。
            try { Console.OutputEncoding = Encoding.UTF8; } catch { }

            // ★只接受**一个**开关。以前 `--selftest --enterprobe` 会静默只跑第一个，
            //   `-enterprobe`（单横线）会直接当成"正常启动"—— 两种都会让人以为
            //   "检查跑过了"，其实什么都没跑。多余 / 认不出的参数一律报错退出。
            // ★v161：`--pooldump <模块名>` 是唯一允许带参数的开关（离线诊断用）。
            // ★`--vbetest <progid>`：直接探某个宿主的 VBE 对象模型（WPS / Word / PPT 兼容性排查）。
            if (args.Length > 1 && !(args.Length == 2 &&
                (args[0] == "--pooldump" || args[0] == "--vbetest" || args[0] == "--declprobe")))
            {
                Console.WriteLine("最多只接受一个开关，这次收到 " + args.Length + " 个：" +
                    string.Join(" ", args));
                return 2;
            }

            if (args.Length >= 1)
            {
                switch (args[0])
                {
                    case "--selftest":    return SelfTest.Run();
                    case "--probe":       return Probe();
                    case "--yieldprobe":  return YieldProbe();
                    case "--writetest":   return WriteTest();
                    case "--uitest":      return UiTest();
                    case "--carettest":   return CaretTest();
                    case "--hktest":      return HkTest();
                    case "--enterprobe":  return EnterProbe();
                    case "--declprobe":   return DeclProbe(args.Length > 1 ? args[1] : null);
                    case "--fmtprobe":    return FmtProbe();
                    case "--formprobe":   return FormProbe();
                    case "--hdrprobe":    return HdrProbe();
                    case "--snipetest":   return SnippetTest();
                    case "--dbtest":      return DbTest();
                    case "--dump":        return DumpModule();
                    // ★v160：诊断用 —— 把当前模块的解析结果 + 光标处的候选全打出来，
                    //   池子里"该有的名字没有"时让用户跑一次，比来回猜快得多。
                    case "--pooldump":    return PoolDump(args.Length > 1 ? args[1] : null);
                    case "--vbetest":     return VbeTest(args.Length > 1 ? args[1] : "Excel.Application");
                    case "--paretest":    return ParenTest();
                    case "--enginetest":  return EngineTest();
                    case "--startuptest": return StartupTest();
                    case "--trusttest":   return TrustTest();
                    // ★v175：正常启动 + 顺带打开设置窗口（桌面快捷方式可以绑这个）。
                    //   不加这个开关的话，"设置"只能从托盘右键进来，而托盘图标
                    //   常被 Win11 收进 ^ 溢出区，用户找不到。
                    case "--settings":    openSettings = true; break;
                }
            }

            // ★认不出的参数不能顺手当成"正常启动"：那会静默起一个托盘实例，
            //   让人以为"检查跑过了"（其实什么都没跑）。打一行用法就退出。
            //   单横线（`-enterprobe`）也算认不出 —— 以前的判据只认 `--`，于是
            //   敲错一个横线就变成"悄悄启动一个常驻进程"。
            // ★v175：`--settings` 是**唯一**不 return 的开关（它要继续正常启动），
            //   所以这里必须把它排除掉 —— 否则 switch 匹配上了、下面照样打用法退出。
            if (args.Length > 0 && !openSettings)
            {
                const string usage =
                    "用法：VbeComplete.exe [开关]\r\n" +
                    "  （不带参数 = 正常启动）\r\n" +
                    "  --selftest    纯逻辑自检（不需要 Excel）\r\n" +
                    "  --dbtest      两个词库的建库/载入检查\r\n" +
                    "  --snipetest   片段库读写检查\r\n" +
                    "  --uitest      候选窗 / 片段窗离屏渲染出 PNG\r\n" +
                    "  --probe       打印当前 VBE 状态（只读）\r\n" +
                    "  --yieldprobe  只采样一次「VBE 自带列表是否可见」（只读）\r\n" +
                    "  --dump        把当前代码窗全文写到 moddump.txt（只读）\r\n" +
                    "  --pooldump    把候选池解析结果 + 光标处候选写到 pooldump.txt（只读）\r\n" +
                    "  --enginetest  用真机模块文本跑候选判定（只读）\r\n" +
                    "  --paretest    真机验证函数补括号（改完即还原）\r\n" +
                    "  --carettest   真机验证中文名光标落点（改完即还原）\r\n" +
                    "  --writetest   真机验证写入链路（改完即还原）\r\n" +
                    "  --hktest      真机验证行操作快捷键（改完即还原）\r\n" +
                    "  --enterprobe  真机验证「回车自动补结构」（走真键盘钩子，改完即还原）\r\n" +
                    "  --declprobe   真机逐字符敲声明位置（Private Fu / ReDim Preserve res），看弹不弹\r\n" +
                    "  --fmtprobe    真机验证「宿主重排空白后光标落在哪儿」（模拟 Ctrl+S 保存的格式化）\r\n" +
                    "  --formprobe   真机验证「窗体代码里提示控件名」（只读，不改模块）\r\n" +
                    "  --startuptest 图标 + 开机自启的登记/回滚（会短暂改注册表，随后还原）\r\n" +
                    "  --trusttest   AccessVBOM 现状（只读；VBECOMPLETE_TRUST_WRITE=1/0 才真写）\r\n" +
                    "  --settings    正常启动，并直接打开设置窗口\r\n";
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
            // --settings：启动后直接把设置窗口摆到前面（顺序必须在 Start 之后 ——
            // 托盘、钩子、词库都要先就绪，设置窗口里那些状态才有意义）
            if (openSettings) app.RunOnUi(() => app.OpenSettings(0));

            // Ctrl+C / 关控制台 时收干净。★必须回 UI 线程：Shutdown 要释放 NotifyIcon
            // 和窗口，从控制台线程直接 Dispose 是跨线程操作，托盘图标会残留。
            Console.CancelKeyPress += (s, e) => { e.Cancel = true; app.RunOnUi(app.Shutdown); };

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
    }
}
