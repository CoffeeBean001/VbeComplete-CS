// COM 接入层：拿到正在运行的 VBE 对象并缓存。
//
// 不用 Microsoft.Office.Interop.*：那套绑定具体 Office 版本，装 WPS / 只有 Word / 换版本
// 都会在编译期绑死。这里全部走 IDispatch 后期绑定（dynamic），与 PY 的 comtypes 同路。
// ★三条硬规矩：① 只有 UI 线程（STA）能调，钩子线程一次都不许碰；② 宿主不在时立刻退让并
//   退避，绝不反复试（宿主忙时失败调用会阻塞到超时，"卡死"就是这么来的）；③ 缓存代理对象，
//   每次用前探活，探活失败立刻丢弃。

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace VbeComplete
{
    internal static class ComCtx
    {
        // ---- oleaut32!GetActiveObject（.NET 5+ 把这个 API 从 Marshal 上删了） ----
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int CLSIDFromProgID(string progId, out Guid pclsid);

        /// <summary>按 ProgID 取正在运行的实例；没有实例返回 null（绝不启动新实例）。</summary>
        public static object GetActive(string progId)
        {
            Guid clsid;
            if (CLSIDFromProgID(progId, out clsid) != 0) return null;
            try
            {
                object o;
                GetActiveObject(ref clsid, IntPtr.Zero, out o);
                return o;
            }
            catch (Exception)
            {
                // MK_E_UNAVAILABLE(0x800401E3) = 宿主没在跑，这是正常情况，不是错误
                return null;
            }
        }

        // ------------------------------------------------------------------
        // 宿主探测
        //
        // ★WPS 12 把文字/表格/演示全塞进同一个 wps.exe —— et.exe / wpp.exe 只是转发器，
        //   把文档交给 wps.exe 后自己立刻退出（所以"按进程名 et 找表格"永远不成立）。
        //   而且 WPS 表格实例在 ROT 里同时注册成 KET.Application 和 Excel.Application。
        //   ⇒ 不猜"一个"ProgID，改成**候选链**：按前台进程名决定先试谁，谁真的有实例、
        //     且 VBE 拿得到，就用谁。先试 KET 是为了避免"Excel 和 WPS 同时开着时连错宿主"。
        // ------------------------------------------------------------------

        private static string ForegroundProcess()
        {
            try { return Native.ProcessNameOfPid(Native.PidOf(Native.GetForegroundWindow())) ?? ""; }
            catch { return ""; }
        }

        private static bool IsWpsProcess(string proc) =>
            string.Equals(proc, "wps", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(proc, "et", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(proc, "wpp", StringComparison.OrdinalIgnoreCase);

        private static bool IsOfficeProcess(string proc) =>
            string.Equals(proc, "EXCEL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(proc, "WINWORD", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(proc, "POWERPNT", StringComparison.OrdinalIgnoreCase);

        private static bool IsHostProcess(string proc) => IsWpsProcess(proc) || IsOfficeProcess(proc);

        /// <summary>
        /// 连到的 Application 是不是"前台那个宿主"。
        /// ★Excel 和 WPS 同时开着时，ROT 里两个 Application 都在，光按候选顺序试会连错 ——
        ///   连错的后果是补全读的是另一个工程的代码（甚至把候选写进另一个宿主）。
        ///   用 Application.Hwnd 的进程号跟前台窗口的进程号比，最省事也最准。
        /// </summary>
        private static bool HostPidMatches(object app, uint fgPid)
        {
            if (fgPid == 0) return true;
            try
            {
                long hw = Convert.ToInt64(((dynamic)app).Hwnd);
                if (hw == 0) return true;
                return Native.PidOf(new IntPtr(hw)) == fgPid;
            }
            catch
            {
                // 取不到 Hwnd（个别宿主/老版本没有这个属性）：别拦着，宁可连上
                return true;
            }
        }

        /// <summary>候选 ProgID（有序）。</summary>
        public static List<string> HostCandidates() => HostCandidates(ForegroundProcess());

        public static List<string> HostCandidates(string proc)
        {
            var list = new List<string>();
            void Add(string id)
            {
                if (!string.IsNullOrEmpty(id) && !list.Contains(id)) list.Add(id);
            }

            if (IsWpsProcess(proc))
            {
                Add("KET.Application");      // WPS 表格
                Add("KWPS.Application");     // WPS 文字
                Add("KWPP.Application");     // WPS 演示
                Add("Excel.Application");    // WPS 表格同时也挂成 Excel.Application
                Add("Word.Application");
                Add("PowerPoint.Application");
                return list;
            }

            // 非 WPS：上次成功的排前面（装多个 Office 时别来回跳）
            if (!string.IsNullOrEmpty(_lastOk)) Add(_lastOk);
            if (string.Equals(proc, "EXCEL", StringComparison.OrdinalIgnoreCase)) Add("Excel.Application");
            else if (string.Equals(proc, "WINWORD", StringComparison.OrdinalIgnoreCase)) Add("Word.Application");
            else if (string.Equals(proc, "POWERPNT", StringComparison.OrdinalIgnoreCase)) Add("PowerPoint.Application");
            Add("Excel.Application");
            Add("Word.Application");
            Add("PowerPoint.Application");
            Add("KET.Application");
            Add("KWPS.Application");
            Add("KWPP.Application");
            return list;
        }

        private static string _lastOk = "";

        public static void NoteHost(string progid) => _lastOk = progid;

        /// <summary>兼容旧调用：猜一个"最可能"的 ProgID（日志 / 探针用）。</summary>
        public static string GuessProgId()
        {
            var list = HostCandidates();
            return list.Count > 0 ? list[0] : "Excel.Application";
        }

        private static object _app;
        private static object _vbe;
        private static DateTime _lastVerify = DateTime.MinValue;

        // 失败退避：同一个前台进程 2 秒内不重试。
        // ★键必须是"前台进程名"而不是"时间"：宿主切换了（用户从 WPS 切到 Excel）要立刻重试，
        //   否则"宿主回来了"会被上一次的退避挡死，表现为永远连不上。
        private static string _failKey;
        private static long _failAt;
        private const long FailBackoffMs = 2000;

        /// <summary>VBE 明明有实例、却拿不到 VBE 对象（最常见原因：AccessVBOM=0）。</summary>
        public static bool VbeAccessDenied { get; private set; }

        // ------------------------------------------------------------------
        // WPS 的引用不能常驻
        //
        // ★实测：只要客户端进程活着并持有 WPS 的 Application 引用，用户关 WPS 时窗口就再也
        //   退不掉（进程还在、Responding=True，等 25 秒也不动）；把客户端杀掉或让进程退出，
        //   立刻 1 秒关掉。Office 不在乎这个，WPS 在乎。
        //   ⇒ WPS 宿主上改成"每轮用完就释放"：PollOnce 开头 Begin()、结尾 End()，
        //     End 时 FinalReleaseComObject。代价是每轮重新 GetActiveObject（ROT 查找，
        //     本地操作，微秒级），换来 WPS 能正常关闭。
        // ------------------------------------------------------------------
        private static bool _wpsHost;
        private static int _depth;
        private static long _lastReleaseAt;
        private static string _loggedProgid = "";
        private static long _loggedAt;

        public static bool IsWpsHost => _wpsHost;

        public static void Begin() => _depth++;

        /// <summary>
        /// 结束一次 COM 使用期。typing=true 表示"用户正在打字"：这时不能释放 ——
        /// 每 80ms 重建一次连接，耗时会盖过快照的 800ms 保鲜期，回车补结构那条链路
        /// 直接失灵（WPS 真机上验到 snap=null）。停手就释放。
        /// ★连续打字超过 2 秒也强制断一次：避免用户"边打字边点关闭"时引用一直挂着。
        /// </summary>
        public static void End(bool typing = false)
        {
            if (_depth > 0) _depth--;
            if (_depth > 0 || !_wpsHost) return;
            if (typing && Environment.TickCount64 - _lastReleaseAt < 2000) return;
            Release();
        }

        /// <summary>释放缓存的对象（WPS 上每轮结束都要调；别的宿主只在出错时调）。</summary>
        public static void Release()
        {
            object vbe = _vbe, app = _app;
            _vbe = null;
            _app = null;
            _lastVerify = DateTime.MinValue;
            _lastReleaseAt = Environment.TickCount64;
            try { if (vbe != null && Marshal.IsComObject(vbe)) Marshal.FinalReleaseComObject(vbe); } catch { }
            try { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); } catch { }
        }

        private static bool IsWpsProgId(string progid) =>
            progid.StartsWith("KET.", StringComparison.OrdinalIgnoreCase) ||
            progid.StartsWith("KWPS.", StringComparison.OrdinalIgnoreCase) ||
            progid.StartsWith("KWPP.", StringComparison.OrdinalIgnoreCase);

        private static string _deniedKey = "";
        private static long _deniedAt;

        /// <summary>丢弃缓存（VBE 关掉了 / 报错之后必须调）。</summary>
        public static void Reset() => Release();

        /// <summary>
        /// 拿 VBE 对象。verify=true 时顺带探一次 ActiveCodePane（更慢但更准）。
        /// 返回 null 表示"现在拿不到"，调用方必须安静退出。
        /// </summary>
        public static dynamic Vbe(bool verify = true)
        {
            if (_vbe != null)
            {
                if (!verify) return _vbe;
                // 200ms 内的重复探活没有意义，直接给
                if ((DateTime.UtcNow - _lastVerify).TotalMilliseconds < 200) return _vbe;
                _lastVerify = DateTime.UtcNow;
                try
                {
                    object cp = ((dynamic)_vbe).ActiveCodePane;
                    if (cp != null) return _vbe;
                }
                catch { }
                Reset();
            }

            string key = ForegroundProcess();
            uint fgPid = Native.PidOf(Native.GetForegroundWindow());
            bool fgIsHost = IsHostProcess(key);
            long now = Environment.TickCount64;
            if (key == _failKey && now - _failAt < FailBackoffMs) return null;

            bool anyHost = false;
            string lastProgid = "";
            Exception lastErr = null;

            foreach (var progid in HostCandidates(key))
            {
                object app = GetActive(progid);
                if (app == null) continue;
                // 前台确实是某个宿主、而这个实例不是它 ⇒ 是另一个宿主，跳过
                if (fgIsHost && !HostPidMatches(app, fgPid)) continue;
                anyHost = true;
                try
                {
                    dynamic dapp = app;
                    dynamic vbe = dapp.VBE;
                    if (ReferenceEquals(vbe, null)) { lastProgid = progid; continue; }

                    _app = app;
                    _vbe = vbe;
                    _lastVerify = DateTime.UtcNow;
                    _failKey = null;
                    VbeAccessDenied = false;
                    _wpsHost = IsWpsProgId(progid);
                    NoteHost(progid);
                    // ★日志限速：WPS 上每轮轮询都要重连（见下面 Release 的说明），
                    //   每轮打一行会把日志和 IO 都拖垮，只在换宿主 / 隔了一分钟时记。
                    long tick = Environment.TickCount64;
                    if (progid != _loggedProgid || tick - _loggedAt > 60000)
                    {
                        _loggedProgid = progid;
                        _loggedAt = tick;
                        Log.Write("COM 已连接：{0}", progid);
                    }
                    return _vbe;
                }
                catch (Exception e)
                {
                    // 有实例、但 VBE 拿不到：多半是"信任对 VBA 工程对象模型的访问"没开
                    lastErr = e;
                    lastProgid = progid;
                    Log.Write("取 VBE 失败（{0}）：{1}", progid, e.Message);
                }
            }

            if (anyHost && lastErr != null && Native.VbeFrameExists())
            {
                // VBE 窗口在、VBE 对象却拿不到 ⇒ 就是"信任对 VBA 工程对象模型的访问"没开
                // （WPS 默认 AccessVBOM=0）。VBE 压根没打开时不是错误，安静等下一次。
                VbeAccessDenied = true;
                NoteDenied(lastProgid);
            }

            _failKey = key;
            _failAt = now;
            return null;
        }

        private static void NoteDenied(string progid)
        {
            // 同样的宿主 30 秒内只记一条，别把日志刷满
            if (progid == _deniedKey && Environment.TickCount64 - _deniedAt < 30000) return;
            _deniedKey = progid;
            _deniedAt = Environment.TickCount64;
            Log.Write("VBA 工程访问被拒：{0}（AccessVBOM={1}）。托盘菜单「允许访问 VBA 工程」可一键开启。",
                progid, VbaTrust.IsBlocked() ? "0" : "非 0");
        }

        public static bool Available => _vbe != null;
    }
}
