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
                // v183：走产品自己那条路径（ProjectScopes -> FormControls）看控件名能不能读到
                sb.AppendLine("--- 组件（只读，走产品路径）---");
                int sig2;
                foreach (var ms in be.ProjectScopes(out sig2))
                    sb.AppendLine("  " + ms.Module + " type=" + ms.ComponentType +
                                  " lines=" + ms.LineCount +
                                  " controls=" + ms.Controls.Count +
                                  (ms.Controls.Count > 0
                                      ? "[" + string.Join(",", ms.Controls) + "]" : ""));
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

        /// <summary>
        /// v183 真机探针：**窗体代码里打控件名**（用户报"我在窗体里，不会提示 TextBox1"）。
        /// <para>★不改模块一个字符、不抢前台：真 Backend（真 Designer 读取 + 真 ProjectScopes）
        /// 配一个合成快照，直接问引擎"这一行该给什么"。验的是这次改动的整条链：
        /// Designer → ModuleScope.Controls → EnsurePool → Trigger → 候选列表。
        /// "打字 → 钩子 → 轮询 → 弹窗"那一段由 declprobe / fmtprobe 各自覆盖，这里不重复。</para>
        /// <para>返回 0 通过 / 1 失败 / 3 SKIP（工程里没有带控件的窗体，一个字都没验到）。</para>
        /// </summary>
        private static int FormProbe()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== formprobe（只读，不改模块） ===");
            int rc = 1;
            try
            {
                var be = new VbeBackend();
                int sig;
                var scopes = be.ProjectScopes(out sig);
                sb.AppendLine("组件（名:类型:控件数） = " + string.Join(", ",
                    scopes.Select(m => m.Module + ":" + m.ComponentType + ":" + m.Controls.Count)));

                ModuleScope form = null, plain = null;
                foreach (var ms in scopes)
                {
                    if (form == null && ms.ComponentType == 3 && ms.Controls.Count > 0) form = ms;
                    if (plain == null && (ms.ComponentType == 1 || ms.ComponentType == 0)) plain = ms;
                }
                // 没有标准模块就退一步用文档模块（ThisWorkbook / Sheet1）做反向对照 ——
                // "别的模块看不见这个窗体的控件名"这条判据与模块种类无关。
                if (plain == null)
                    foreach (var ms in scopes)
                        if (ms.ComponentType != 3) { plain = ms; break; }
                if (form == null)
                {
                    sb.AppendLine("SKIP：这个工程里没有带控件的窗体（UserForm）—— 一个字都没验到。");
                    rc = 3;
                }
                else
                {
                    // 拿一个真控件名、用它的前两个字母当前缀（用户截图里就是 `te` -> TextBox1）
                    string ctl = form.Controls.FirstOrDefault(x => x != null && x.Length >= 2);
                    string prefix = ctl.Substring(0, Math.Min(3, ctl.Length)).ToLowerInvariant();
                    string line = "    s = " + prefix;
                    sb.AppendLine("窗体 = " + form.Module + "  控件 = " +
                                  string.Join(",", form.Controls));
                    sb.AppendLine("用例 = [" + line + "]  -> 期望出现 " + ctl);

                    string ftext = be.ModuleTextOf(form.Module);
                    if (ftext == null)
                    {
                        sb.AppendLine("SKIP：读不到 " + form.Module + " 的全文（工程被锁？）。");
                        rc = 3;
                    }
                    else
                    {
                        rc = 0;
                        // 1) 该弹：窗体的代码里打控件名前缀
                        if (!Check(be, form.Module, ftext, line, ctl, true, sb)) rc = 1;
                        // 2) 反向：同一行放进标准模块，那个控件名不许冒出来
                        if (plain != null)
                        {
                            string ptext = be.ModuleTextOf(plain.Module);
                            if (ptext != null && !Check(be, plain.Module, ptext, line, ctl, false, sb))
                                rc = 1;
                            else if (ptext == null)
                                sb.AppendLine("  （反向对照跳过：读不到 " + plain.Module + " 的全文）");
                        }
                        else sb.AppendLine("  （反向对照跳过：工程里没有标准模块）");
                    }
                }
            }
            catch (Exception e) { sb.AppendLine("EXCEPTION " + e); rc = 1; }
            sb.AppendLine(rc == 0 ? "RESULT: PASS" : rc == 3 ? "RESULT: SKIP" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "formprobe.log"), sb.ToString(), Encoding.UTF8);
            Console.WriteLine(sb.ToString());
            return rc;
        }

        /// <summary>formprobe 的一格：把 <paramref name="line"/> 当成"模块的第 1 行"问引擎。</summary>
        private static bool Check(VbeBackend be, string module, string text, string line,
            string ctl, bool mustHave, StringBuilder sb)
        {
            var eng = new Engine { Backend = new FixedTextBackend(be, module, text) };
            var sn = new Snapshot
            {
                Line = 1,
                Text = line,
                Module = module,
                CaretCol = line.Length + 1,     // 光标在行尾（VBE 的列 = 字符数 + 1）
            };
            bool shown = eng.Trigger(false, sn);
            var names = eng.Items.Select(x => x.Name).ToList();
            bool hit = names.Any(n => string.Equals(n, ctl, StringComparison.OrdinalIgnoreCase));
            bool pass = mustHave ? (shown && hit) : !hit;

            sb.AppendLine((pass ? "  PASS  " : "  FAIL  ") +
                          (mustHave ? "窗体里打前缀出控件名" : "控件名不跨模块泄露") +
                          "（" + module + "）");
            sb.AppendLine("         弹=" + shown + " 条数=" + eng.Items.Count +
                          " 候选=[" + string.Join(",", names.Take(24)) + "]");
            if (mustHave && !shown)
            {
                // 没弹多半是"打全即收"（只剩一条候选就是它自己），换手动触发再问一次 ——
                // 这样 FAIL 时能一眼分清是位置判据挡了、还是被 HIDE_ON_COMPLETE 收了。
                bool man = eng.Trigger(true, sn);
                sb.AppendLine("         手动触发=" + man + " 候选=[" +
                    string.Join(",", eng.Items.Select(x => x.Name).Take(24)) + "]");
            }
            eng.Hide();
            return pass;
        }

        /// <summary>
        /// 把"某个模块的全文"钉住的后端：其余能力全转给真 <see cref="VbeBackend"/>。
        /// 只为让探针能在**不切换代码窗**的前提下问"另一个模块里该给什么候选"。
        /// </summary>
        private sealed class FixedTextBackend : IVbeSource
        {
            private readonly VbeBackend _be;
            private readonly string _module, _text;
            public FixedTextBackend(VbeBackend be, string module, string text)
            { _be = be; _module = module; _text = text; }

            public Snapshot Snapshot(bool verify = true) => null;
            public string ModuleText() => _text;
            public List<string> ComponentNames() => _be.ComponentNames();
            public List<ModuleScope> ProjectScopes(out int signature) =>
                _be.ProjectScopes(out signature);
            public bool ApplyCompletion(int lineNo, int wordStartCol, int endCol,
                string completion, bool wantParens, bool wantSpace, string enclosingProc) => false;
            public string EnclosingProcName() => Parser.ProcOfLine(
                Scope.Analyze(_module, _text).Procs, 1)?.Name;
        }
    }
}
