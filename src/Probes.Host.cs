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
    /// <summary>宿主探针：能不能连上 VBE、按键模拟是否生效。</summary>
    internal static partial class Program
    {
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
                            // CodePane 没有 Selection 属性，GetSelection 又是 4 个 ByRef
                            // （dynamic 填不回来）—— 走产品自己那条 ComCall.CallFourOut。
                            if (be.TryGetSelection(out int sl, out int sc, out int el, out int ec))
                                sb.AppendLine("  GetSelection = " + sl + "," + sc + " .. " + el + "," + ec);
                            else
                                sb.AppendLine("  GetSelection = 失败（没有活动代码窗，正常）");
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
        /// 直接探一个 ProgID 的 VBE 对象模型（WPS / Word / PPT 兼容性排查）。
        /// ★每一步单独 try：坏在"拿不到 VBE"和坏在"拿不到 CodeModule"是完全不同的两件事，
        ///   合成一个 catch 只会看到一句"发生意外"。
        /// </summary>
        private static int VbeTest(string progid)
        {
            var sb = new StringBuilder();
            sb.AppendLine("progid  = " + progid);
            sb.AppendLine("64bit   = " + Environment.Is64BitProcess);

            object app = ComCtx.GetActive(progid);
            sb.AppendLine("GetActive = " + (app == null ? "null（ROT 里没有这个实例）" : "ok"));
            if (app != null)
            {
                dynamic dapp = app;
                try { sb.AppendLine("  Name    = " + Convert.ToString(dapp.Name)); }
                catch (Exception e) { sb.AppendLine("  Name    ERR " + e.Message); }
                try { sb.AppendLine("  Version = " + Convert.ToString(dapp.Version)); }
                catch (Exception e) { sb.AppendLine("  Version ERR " + e.Message); }
                try
                {
                    dynamic books = dapp.Workbooks;
                    int bn = Convert.ToInt32(books.Count);
                    var names = new List<string>();
                    for (int i = 1; i <= bn; i++) names.Add(Convert.ToString(books.Item(i).Name));
                    sb.AppendLine("  Workbooks = " + bn + " [" + string.Join(", ", names) + "]");
                }
                catch (Exception e) { sb.AppendLine("  Workbooks ERR " + e.Message); }
                try
                {
                    dynamic vbe = dapp.VBE;
                    sb.AppendLine("  VBE     = " + (ReferenceEquals(vbe, null) ? "null" : "ok"));
                    if (vbe != null)
                    {
                        try
                        {
                            dynamic mw = vbe.MainWindow;
                            sb.AppendLine("  MainWindow = " + Convert.ToString(mw.Caption) +
                                          " hwnd=" + Convert.ToInt32(mw.HWND));
                        }
                        catch (Exception e) { sb.AppendLine("  MainWindow ERR " + e.Message); }
                        try
                        {
                            dynamic ps = vbe.VBProjects;
                            int n = Convert.ToInt32(ps.Count);
                            sb.AppendLine("  VBProjects = " + n);
                            for (int i = 1; i <= n; i++)
                            {
                                dynamic p = ps.Item(i);
                                sb.AppendLine("    [" + i + "] " + Convert.ToString(p.Name));
                            }
                        }
                        catch (Exception e) { sb.AppendLine("  VBProjects ERR " + e.Message); }
                        try
                        {
                            dynamic cp = vbe.ActiveCodePane;
                            sb.AppendLine("  ActiveCodePane = " + (ReferenceEquals(cp, null) ? "null" : "ok"));
                            if (cp != null)
                            {
                                dynamic cm = cp.CodeModule;
                                int cnt = Convert.ToInt32(cm.CountOfLines);
                                sb.AppendLine("    module = " + Convert.ToString(cm.Name) + " lines = " + cnt);
                                if (cnt > 0)
                                {
                                    int take = Math.Min(cnt, 3);
                                    string txt = Convert.ToString(cm.Lines(1, take)) ?? "";
                                    sb.AppendLine("    first " + take + " = " +
                                        txt.Replace("\r", "\\r").Replace("\n", "\\n"));
                                }
                            }
                        }
                        catch (Exception e) { sb.AppendLine("  ActiveCodePane ERR " + e.Message); }
                    }
                }
                catch (Exception e) { sb.AppendLine("  VBE     ERR " + e.Message); }
            }

            string text = sb.ToString();
            File.WriteAllText(Path.Combine(Settings.Home, "vbetest.log"), text, Encoding.UTF8);
            Console.WriteLine(text);

            // 诊断用：VBECOMPLETE_HOLD_SEC=30 让本进程读完不退出（保持 COM 引用），
            // 用来验证"客户端一直活着"是不是宿主关不掉的原因。
            int hold = 0;
            int.TryParse(Environment.GetEnvironmentVariable("VBECOMPLETE_HOLD_SEC") ?? "", out hold);
            if (hold > 0) Thread.Sleep(hold * 1000);

            return app == null ? 1 : 0;
        }

        /// <summary>
        /// 让位判据采样：先把 VBE 自己的成员列表**真弹出来**（抢前台 + Ctrl+Space），
        /// 再采样 App.PollOnce 用的同一个函数。
        /// ★不促发的话这条判据永远读到 false —— 那是"绿着但什么都没验"。
        /// </summary>
        private static int YieldProbe()
        {
            var sb = new StringBuilder();
            int visible = 0, seen = 0;
            try
            {
                var be = new VbeBackend();
                Snapshot snap = be.Snapshot();
                if (snap == null)
                {
                    sb.AppendLine("no code module, skip");
                    File.WriteAllText(Path.Combine(Settings.Home, "yieldprobe.log"), sb.ToString(), Encoding.UTF8);
                    return 1;
                }
                sb.AppendLine("module = " + snap.Module + "  caret = L" + snap.Line + " C" + snap.CaretCol);

                IntPtr frame = Native.FindWindow("wndclass_desked_gsk", null);
                bool fg = false;
                for (int i = 0; i < 10 && !fg; i++)
                {
                    Native.ForceForeground(frame);
                    be.SetCaret(snap.Line, snap.CaretCol);
                    Thread.Sleep(150);
                    fg = frame != IntPtr.Zero && Native.GetForegroundWindow() == frame;
                }
                uint tid = 0;
                if (frame != IntPtr.Zero)
                {
                    uint pid;
                    tid = Native.GetWindowThreadProcessId(frame, out pid);   // ★返回值才是线程 id
                }
                sb.AppendLine("foreground(VBE) = " + fg + "  focusClass = " +
                              Native.ClassNameOf(Native.FocusedHwndOf(tid)));
                if (!fg)
                {
                    sb.AppendLine("抢不到前台，促发不了列表（结论无效）");
                    File.WriteAllText(Path.Combine(Settings.Home, "yieldprobe.log"), sb.ToString(), Encoding.UTF8);
                    return 1;
                }

                Native.SendVk(Native.VK_SPACE, ctrl: true);
                Thread.Sleep(400);
            }
            catch (Exception e) { sb.AppendLine("促发异常 " + e.Message); }

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
            // 探针不许留下副作用：把促发出来的列表按 Esc 关掉
            if (visible > 0) { Native.SendVk(Native.VK_ESCAPE); Thread.Sleep(120); }
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
