// 主流程：全局钩子 + 自适应轮询 + 按键路由 + "改一行"类操作。
// ★这是唯一有权决定吃不吃一个键的地方：钩子线程只读内存状态（一次 COM 都不许碰），
//   真动作投递到 UI 线程；吃掉却没做成事，必须原样还回去。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace VbeComplete
{
    internal sealed class App : ApplicationContext
    {
        private const string CodePaneClass = "VbaWindow";
        private const string VbeFrameClass = "wndclass_desked_gsk";
        private static readonly string[] VbePopupClasses = { "NameListWndClass", "PopupTipWndClass" };

        private readonly VbeBackend _be = new VbeBackend();
        private readonly Recent _recent = new Recent();
        private long _recentFlushAt;                      // 上次把 recent.json 写盘的时刻
        private readonly Engine _eng;
        private readonly PopupForm _popup = new PopupForm();
        private readonly Control _disp = new Control();

        private System.Windows.Forms.Timer _timer;
        private NotifyIcon _tray;
        private EventWaitHandle _showEvent;
        private Thread _showThread;

        // ---- 跨线程共享的状态（钩子线程只读，UI 线程写） ----
        private volatile bool _codePaneFocused;
        private volatile bool _popupVisible;
        private long _lastYieldLog;               // "让位"日志限速（每 2 秒最多一条）
        private long _lastJumpLog;                // "跳行不弹"日志限速（每 2 秒最多一条）
        private volatile bool _running = true;
        private volatile Snapshot _hookSnap;      // 最近一次轮询拿到的快照（钩子线程只读）
        private long _popupBeat;                  // 候选窗心跳（UI 线程写，钩子线程只读）
        private volatile int _pendingKeys;        // 按了"可能写字"的键、文档却没变（输入法组字）
        private long _pendingAt;                  // 上面那次计数的时刻（超时就当 0）
        private long _hookSnapTicks;
        // 片段窗的 HWND。★必须由 UI 线程在 Show 之前存好、钩子线程只读这个
        // IntPtr：钩子线程去碰 `_snippetForm.Handle` 会在句柄还没建出来时
        // 触发一次【在钩子线程上创建窗口句柄】，那是跨线程建窗，必炸。
        private volatile IntPtr _snippetHwnd;
        private long _lastKeyTs;
        // ★v179：最近一次"能改空白"的按键（空格 / Tab / 退格 / Delete）的时刻。
        //   宿主重排空白后的"光标搬家"要看它 —— 用户自己在敲空格/退格时绝不能动光标。
        //   ★v174 用的是"任何编辑键"，于是"敲完 `1` 立刻按 Ctrl+S"（间隔常 < 400ms）
        //     被当成人手改空白，搬家整条被跳过 = 用户报的"错位又复现"。
        private long _lastBlankTs;
        // ★v179：最近一次 Ctrl+S 的时刻。VBE 保存时会顺手把当前行格式化，
        //   这是比"数空格"更硬的前兆 —— 刚按过保存，空白怎么变都是它干的。
        private long _lastSaveTs;

        // ---- 探针专用：0 = 注入键一律不理（= 生产行为），-1 = 全放行，其它 = 只放行这一个 vk ----
        // 钩子必须跳过 LLKHF_INJECTED（否则自己补的空格/括号会自己触发自己），而
        // --enterprobe 要验的正是这条路径，所以留了这个只对测试键生效的放行口。
        internal static volatile int AcceptInjectedVk;

        // ---- 探针用的只读视图（生产路径一个都不读） ----
        internal bool CodePaneFocused => _codePaneFocused;
        internal bool PopupVisible => _popupVisible;
        /// <summary>候选窗里此刻排在最前面的几个名字（没弹就是空串），只给探针看。</summary>
        internal string PopupNames
        {
            get
            {
                if (!_popupVisible) return "";
                var names = new List<string>();
                var items = _eng.Items;
                for (int i = 0; i < items.Count && i < 8; i++) names.Add(items[i].Name);
                return string.Join(",", names);
            }
        }
        internal int PendingKeys => _pendingKeys;
        internal Snapshot HookSnap => _hookSnap;
        internal long HookSnapAgeMs => Environment.TickCount64 - _hookSnapTicks;
        internal void RunOnUi(Action a) { try { _disp.Invoke(a); } catch { } }
        internal void RequestExit() { try { _disp.BeginInvoke((Action)ExitThread); } catch { } }

        // ---- 钩子 ----
        private Thread _hookThread;
        private IntPtr _hookHandle = IntPtr.Zero;
        private uint _hookThreadId;
        private Native.LowLevelKeyboardProc _hookProc;   // 必须保持引用，否则被 GC
        // 已吞掉 keydown 的键 -> 吞键时刻。★必须是"带时标"的字典而不是 HashSet：
        // 吞了 keydown 之后如果那个 keyup 没回来（切窗口、被别的钩子劫走、程序被挂起），
        // 键就永远留在表里，下一次同键的 keyup 会在**任何**窗口里被吞掉 —— 丢一个
        // keyup 的后果是修饰键粘住（Shift 按一下就一直是选定状态）。所以只认
        // 距今 1.5 秒内的配对，超时当没这回事。
        private readonly Dictionary<int, long> _swallowed = new Dictionary<int, long>();
        private const long SwallowPairMs = 1500;

        // ---- 轮询状态 ----
        private Snapshot _lastSnap;
        private string _lastModName = "";
        private bool _deniedShown;                // "VBE 访问被拒"的气泡只弹一次
        private int _shutdownDone;                // Shutdown 的幂等标志（Interlocked）

        // ---- 代码片段 ----
        private SnippetStore _snippetStore;
        private SnippetForm _snippetForm;

        // ---- 设置界面（托盘「设置…」/ 双击图标 / 「添加关键字…」共用这一个实例） ----
        private SettingsForm _settingsForm;

        // ★v175：这一组不再是 readonly 快照 —— 设置界面保存后由 ReloadSettings() 重读。
        //   钩子线程只读它们（volatile 那个另有说明），UI 线程写，所以是普通字段。
        private bool NoCtrlNav;
        private bool NoSnippets;
        private bool NoComment;
        private bool NoDupLine;
        private bool NoDelLine;
        private bool NoInsertAbove;
        private bool NoEnterIndent;
        private bool NoAutoClose;
        private bool NoBranchAlign;
        private bool NoShiftNav;
        private bool NoAutoPair;
        private bool NoBsDelLine;
        private bool NoBsDelPair;
        // 自动弹窗只认按键：鼠标右键粘贴 / 拖拽 / 别的程序改代码造成的文本变化
        // 只收起、不弹（Python 版 REQUIRE_KEY_TRIGGER，默认开）。
        private bool RequireKeyTrigger = true;
        /// <summary>
        /// v174 的口子：关掉"宿主重排空白"的处理。只为 --fmtprobe 的 A/B 对照存在 ——
        /// 探针得能证明"把这段关掉就复现了用户报的现象"，否则绿了也不知道绿的是什么。
        /// </summary>
        private readonly bool NoRespacedFix = Settings.Flag("NO_RESPACED_FIX", false);
        private const long KeyTriggerWindowMs = 1200;
        /// <summary>
        /// "最近这么久没按过空白键"才认"空白是怎么变的不是用户干的"（见 PollOnce 里的
        /// hostRespaced）。打字时轮询 50ms 一跳，人敲一个空格不可能漏过这个窗口。
        /// ★v179：v174 之前这里是 400ms + "任何编辑键"，把"敲完立刻保存"挡在了门外。
        /// </summary>
        private const long NoBlankKeyMs = 300;
        /// <summary>刚按过 Ctrl+S 的认账窗口 —— VBE 保存时格式化当前行就在这几百毫秒内。</summary>
        private const long SaveWindowMs = 1200;

        public App()
        {
            // 强制建句柄：跨线程 BeginInvoke 需要一个真窗口
            var _ = _disp.Handle;

            _be.SendChar = ch => Native.SendUnicodeChar(ch);
            _be.SendVk = (vk, ctrl, shift) => Native.SendVk(vk, ctrl, shift);

            // ★词库必须在这里、在引擎干活之前载入：两个 SQLite 库（关键字 / 函数）
            //   一次性读进内存，之后每一次按键都是纯内存匹配。
            new WordLibrary().Load();

            _eng = new Engine
            {
                Backend = _be,
                Recent = _recent,
            };
            _eng.OnShow = (items, sel, top) =>
            {
                _popup.SetData(items, sel, top, _eng.ViewRows);
                ShowPopupAtCaret();
                // ★镜像标志必须跟着【真实可见性】走：ShowPopupAtCaret 拿不到插入符
                //   时会 _eng.Hide()，此时窗口压根没显示。无条件置 true 会让
                //   _popupVisible 永久为真 —— 之后 ↑/↓/PgUp/Esc 全被吞掉却什么都
                //   不发生（且 Engine.Hide 的 `if (!Visible) return` 让 OnHide
                //   再也纠正不了它）。
                _popupVisible = _popup.Visible;
                _popupBeat = Environment.TickCount64;
                if (!_popupVisible) _eng.Hide();
            };
            _eng.OnHide = () =>
            {
                _popupVisible = false;
                _popup.ResetPlacement();
                if (_popup.Visible) _popup.Hide();
            };
            _eng.OnMove = () =>
            {
                _popup.UpdateView(_eng.SelectedIndex, _eng.TopIndex);
            };

            _popup.RowClicked += i => { _eng.SetSelected(i); _eng.Accept(); };
            _popup.Wheel += lines => _eng.ScrollBy(lines);
            _popup.ScrollTo += top =>
            {
                _eng.SetTop(top);
                _popup.UpdateView(_eng.SelectedIndex, _eng.TopIndex);
            };
            // VBE 自己的成员列表弹出来的时候我们让位（它的类型信息比我们准）
            _eng.YieldToVbe = () => Native.VbeMemberListVisible();

            // ★v175：这些开关以前靠字段初始化器读一次，现在改成可重载 ——
            //   构造时先灌一遍，之后设置界面保存会再灌（ReloadSettings）。
            ReloadSettings();
        }

        /// <summary>
        /// 重新读一遍设置（保存 / 恢复默认之后调）：钩子侧开关字段、候选窗字体配色、候选池。
        /// ★任何一项出错都不能拦住别的项 —— 点保存不许因为某个分支异常而"看起来什么都没发生"。
        /// </summary>
        internal void ReloadSettings()
        {
            NoCtrlNav = Settings.Flag("NO_CTRL_NAV", false);
            NoSnippets = Settings.Flag("NO_SNIPPETS", false);
            NoComment = Settings.Flag("NO_COMMENT", false);
            NoDupLine = Settings.Flag("NO_DUP_LINE", false);
            NoDelLine = Settings.Flag("NO_DEL_LINE", false);
            NoInsertAbove = Settings.Flag("NO_INSERT_ABOVE", false);
            NoEnterIndent = Settings.Flag("NO_ENTER_INDENT", false);
            NoAutoClose = Settings.Flag("NO_AUTOCLOSE", false);
            NoBranchAlign = Settings.Flag("NO_BRANCH_ALIGN", false);
            NoShiftNav = Settings.Flag("NO_SHIFT_NAV", false);
            NoAutoPair = Settings.Flag("NO_AUTOPAIR", false);
            NoBsDelLine = Settings.Flag("NO_BS_DELLINE", false);
            NoBsDelPair = Settings.Flag("NO_BS_DELPAIR", false);
            RequireKeyTrigger = Settings.Flag("REQUIRE_KEY_TRIGGER", true);
            try { _eng.MemberPopup = Settings.Flag("MEMBER_POPUP", false); } catch { }
            try { _popup.ReloadStyle(); } catch (Exception e) { Log.Write("候选窗重读外观失败：{0}", e.Message); }
            try { _eng.InvalidatePool(); } catch { }
            Log.Write("设置已重载：Ctrl+Nav={0} 片段={1} 注释={2} 按键触发={3}",
                !NoCtrlNav, !NoSnippets, !NoComment, RequireKeyTrigger);
        }

        /// <summary>打开设置窗口（page：0 常用 / 1 快捷键 / 2 候选 / 3 外观 / 4 词库）。</summary>
        internal void OpenSettings(int page = 0)
        {
            try
            {
                _eng.Hide();
                if (_settingsForm == null || _settingsForm.IsDisposed)
                {
                    _settingsForm = new SettingsForm(ReloadSettings, () => _eng.InvalidatePool(), page);
                    _settingsForm.FormClosed += (s, e) =>
                    {
                        _settingsForm = null;
                        try { _be.FocusVbe(); } catch { }
                    };
                }
                else _settingsForm.GoPage(page);

                if (_settingsForm.WindowState == FormWindowState.Minimized)
                    _settingsForm.WindowState = FormWindowState.Normal;
                _settingsForm.Show();
                _settingsForm.Activate();
                _settingsForm.BringToFront();
            }
            catch (Exception e) { Log.Write("打开设置窗口失败：{0}", e); }
        }

        public void Start()
        {
            // 开机自启：按 autostart.txt 里记的意图把注册表修正一遍（第一次运行默认开）。
            // ★必须放在 SetupTray 之前：托盘菜单里的勾选状态是在建菜单时读注册表的，
            //   顺序反过来会看到"已经登记了但菜单上没勾"。
            try { AutoStart.Sync(); } catch (Exception e) { Log.Write("开机自启同步失败：{0}", e.Message); }
            SetupTray();
            StartHook();
            StartShowListener();
            _timer = new System.Windows.Forms.Timer { Interval = 200 };
            _timer.Tick += (s, e) => PollOnce();
            _timer.Start();
            Log.Write("C# 版已启动（Home={0}，{1}）", Settings.Home, AutoStart.Describe());

            // 托盘图标刚注册时，Windows 才刚建好它在 NotifyIconSettings 里的那条记录，
            // 所以延后 2.5 秒再去找 —— 太早会找不到（第一次运行尤其如此）。
            var pin = new System.Windows.Forms.Timer { Interval = 2500 };
            pin.Tick += (s, e) =>
            {
                pin.Stop();
                pin.Dispose();
                try { TrayPin.EnsureVisible(); } catch { }
            };
            pin.Start();

            // 启动了就必须让用户看得见。这个程序没有主窗口，图标又默认折叠进
            // Win11 的托盘溢出区 —— 不给提示的话，"双击了没反应"和"已经正常运行"
            // 在屏幕上长得一模一样（用户就是这么报的 bug）。
            // ★只在第一次运行时提示：现在它会开机自启，每次登录都弹一个气泡就是骚扰。
            //   之后手动再启动时，若已有一份在跑，SignalExisting() 会负责提示。
            string first = Path.Combine(Settings.Home, "firstrun.txt");
            if (!File.Exists(first))
            {
                ShowBalloon("VbeComplete 已启动",
                    "没有主窗口：图标在右下角托盘里（若看不到，点一下任务栏的 ^ ，把它拖到外面即可常驻）。\r\n"
                    + "右键那个图标可以开关「开机自动启动」或退出。");
                try { File.WriteAllText(first, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")); } catch { }
            }
        }

        private void ShowBalloon(string title, string text)
        {
            try
            {
                if (_tray == null) return;
                _tray.BalloonTipTitle = title;
                _tray.BalloonTipText = text;
                _tray.BalloonTipIcon = ToolTipIcon.Info;
                _tray.ShowBalloonTip(5000);
            }
            catch (Exception e) { Log.Write("气泡提示失败：{0}", e.Message); }
        }

        /// <summary>
        /// 托盘「允许访问 VBA 工程」点完之后给用户看的那段话。
        /// ★抽成静态方法是为了让 `--trusttest` 弹**同一份**文字：用户将看到的字，和我验过的
        ///   字必须是同一份，否则"验过了"是空话。
        /// </summary>
        internal static string TrustMessage(bool wasOn, int n)
        {
            var msg = new StringBuilder();
            if (wasOn)
            {
                msg.Append(n > 0 ? "已关闭 " + n + " 处。\r\n\r\n" : "本来就没开着。\r\n\r\n");
                msg.Append("关掉之后，工具在 WPS / Excel 里就拿不到 VBA 工程 —— "
                         + "补全、回车补结构、Ctrl+/ 都会失效。\r\n");
                msg.Append("想恢复：关掉这个框，再点一次这一项（勾会重新出现）。\r\n\r\n");
            }
            else
            {
                msg.Append(n > 0 ? "已开启（写入 " + n + " 处）。\r\n\r\n"
                                 : "12 处注册表本来就是开着的，这次没有改动。\r\n\r\n");
            }
            msg.Append("当前状态（AccessVBOM）：\r\n");
            msg.Append(VbaTrust.Detail());
            msg.Append("\r\n");
            msg.Append(n > 0
                // 只在宿主启动时读一次 ⇒ 不重开就等于没改
                ? "要让它在 WPS / Excel 里生效，\r\n必须把 WPS / Excel 完全退出（连托盘图标一起）再重新打开。\r\n"
                : "如果 WPS 里还是没反应，多半不是这个开关的问题。\r\n");
            if (!wasOn)
                msg.Append("另外：WPS 必须先按一次 Alt+F11 打开过 VBE 窗口，工具才连得上。");
            return msg.ToString();
        }

        /// <summary>
        /// 监听"又有人启动了第二个实例"。Program 那边发现互斥体已存在时会 Set
        /// 这个事件，我们收到就把托盘图标的存在感刷一遍 —— 否则用户双击后的
        /// 全部观感就是"什么都没发生"。
        /// </summary>
        private void StartShowListener()
        {
            try
            {
                _showEvent = new EventWaitHandle(false, EventResetMode.AutoReset, Program.ShowEventName);
                _showThread = new Thread(() =>
                {
                    while (_running)
                    {
                        try
                        {
                            if (!_showEvent.WaitOne(500)) continue;
                            if (!_running) return;
                            // 必须回 UI 线程碰 NotifyIcon
                            if (_disp.InvokeRequired)
                                _disp.BeginInvoke((Action)(() => ShowBalloon(
                                    "VbeComplete 已经在运行",
                                    "不用再启动一次。图标在右下角托盘的溢出区里，右键可以退出。")));
                            else
                                ShowBalloon("VbeComplete 已经在运行",
                                    "不用再启动一次。图标在右下角托盘的溢出区里，右键可以退出。");
                        }
                        catch { return; }
                    }
                })
                { IsBackground = true, Name = "show-listener" };
                _showThread.Start();
            }
            catch (Exception e) { Log.Write("重复启动监听未装上：{0}", e.Message); }
        }

        /// <summary>
        /// 托盘图标：没有它，用户只能去任务管理器杀进程。
        /// ★图标必须用【系统自带】的 —— 不引入任何资源文件，保持单文件干净。
        /// </summary>
        private void SetupTray()
        {
            try
            {
                var menu = new ContextMenuStrip();
                var title = new ToolStripMenuItem("VbeComplete · VBA 补全（C# 版）") { Enabled = false };
                menu.Items.Add(title);
                menu.Items.Add(new ToolStripSeparator());
                // ★v175：设置界面。以前改一项配置要拿记事本翻 settings.ini 找那一行、
                //   删掉行首分号；两个词库更是得装 DB Browser 才能加一个字。
                menu.Items.Add("设置…", null, (s, e) => OpenSettings(0));
                // 单独挂一条"添加关键字"：这是用户最常来的地方，埋在设置的第二层
                // 就等于没做（他原来的办法是"我不知道怎么操作数据库"）。
                menu.Items.Add("添加关键字 / 函数…", null, (s, e) => OpenSettings(4));
                menu.Items.Add("代码片段…", null, (s, e) =>
                {
                    try { ToggleSnippets(); }
                    catch (Exception ex) { Log.Write("托盘打开片段窗失败：{0}", ex.Message); }
                });

                // 开机自启：勾选状态按注册表的**真实情况**显示（不是按配置），
                // 免得用户看到"勾着但其实早被清理软件删了"。
                var miAutoStart = new ToolStripMenuItem("开机自动启动")
                {
                    CheckOnClick = false,                    // 自己处理，避免和真实状态脱节
                    Checked = AutoStart.IsRegistered(out _),
                    ToolTipText = "随 Windows 登录自动运行（写 HKCU 的 Run 键，不需要管理员）",
                };
                miAutoStart.Click += (s, e) =>
                {
                    bool want = !AutoStart.IsRegistered(out _);
                    bool ok = AutoStart.SetEnabled(want);
                    miAutoStart.Checked = AutoStart.IsRegistered(out _);
                    Log.Write("托盘切换开机自启 -> {0}：{1}", want ? "开" : "关", ok ? "成功" : "失败");
                    if (!ok)
                        ShowBalloon("开机自启没设置成功",
                            "写注册表被拒绝了（常见于安全软件拦截）。日志里有详细原因。");
                };
                menu.Items.Add(miAutoStart);
                // 允许访问 VBA 工程：点一下开、再点一下关（v167）。
                // ★这一项原来只弹气泡，Win11 上气泡静默不显示 ⇒ 用户点了"完全没反应"
                //   （实测反馈）。改注册表这种动作必须让用户当场看到结果，所以走模态框。
                var miTrust = new ToolStripMenuItem("允许访问 VBA 工程（WPS / Office）")
                {
                    ToolTipText = "点一下切换开 / 关（勾 = 已开）。写 HKCU 的 AccessVBOM，不需要管理员。"
                                + "关着时工具在 WPS 里拿不到 VBE。",
                };
                miTrust.Click += (s, e) =>
                {
                    // ★WPS 默认 AccessVBOM=0，此时 Application.VBE 直接 E_FAIL —— 关着这个
                    //   开关，工具在 WPS 里就像"完全没反应"。
                    bool wasOn = VbaTrust.AllEnabled();
                    int n = wasOn ? VbaTrust.DisableAll() : VbaTrust.EnableAll();
                    miTrust.Checked = VbaTrust.AllEnabled();
                    Log.Write("托盘切换 AccessVBOM -> {0}：改了 {1} 处（{2}）",
                        wasOn ? "关" : "开", n, VbaTrust.Describe());
                    InfoBox.Show("允许访问 VBA 工程", TrustMessage(wasOn, n));
                };
                menu.Items.Add(miTrust);
                // 每次展开菜单都按注册表刷新勾选 —— 用户可能手工在 Office 里改过
                menu.Opening += (s, e) =>
                {
                    miTrust.Checked = VbaTrust.AllEnabled();
                    miAutoStart.Checked = AutoStart.IsRegistered(out _);
                };
                menu.Items.Add("打开日志", null, (s, e) =>
                {
                    // ★v151：路径必须加引号 —— 家目录带空格（`D:\我的 工具\…`）时
                    //   记事本会把路径切成两段，这条菜单恰恰是用户排查问题的唯一入口。
                    try
                    {
                        var psi = new System.Diagnostics.ProcessStartInfo(
                            "notepad.exe", "\"" + Log.Path_ + "\"") { UseShellExecute = true };
                        System.Diagnostics.Process.Start(psi);
                    }
                    catch (Exception ex) { Log.Write("打开日志失败：{0}", ex.Message); }
                });
                menu.Items.Add("退出", null, (s, e) => Shutdown());

                _tray = new NotifyIcon
                {
                    // ★一定要用自己的图标：SystemIcons.Application 是一张通用的
                    //   "空白窗口"，在托盘里跟十几个程序长得一模一样，用户根本找不到。
                    Icon = IconAssets.Tray(),
                    Text = "VbeComplete · VBA 补全（C# 版）",
                    Visible = true,
                    ContextMenuStrip = menu,
                };
                // 双击图标 = 打开设置（托盘程序的惯例，也是用户第一次点进来最想找的东西）
                _tray.DoubleClick += (s, e) =>
                {
                    var snap = _be.Snapshot();
                    Log.Write("托盘双击：snapshot={0}", snap == null ? "null" : snap.ToString());
                    OpenSettings(0);
                };
            }
            catch (Exception e) { Log.Write("托盘创建失败：{0}", e.Message); }
        }

        // ------------------------------------------------------------------
        // 钩子线程
        // ------------------------------------------------------------------

        private void StartHook()
        {
            _hookThread = new Thread(HookLoop)
            {
                IsBackground = true,
                Name = "kbd-hook",
            };
            _hookThread.Start();
        }

        private void HookLoop()
        {
            _hookThreadId = Native.GetCurrentThreadId();
            _hookProc = HookProc;
            IntPtr mod = Native.GetModuleHandle(null);
            _hookHandle = Native.SetWindowsHookEx(Native.WH_KEYBOARD_LL, _hookProc, mod, 0);
            if (_hookHandle == IntPtr.Zero)
            {
                Log.Write("装钩子失败，错误码 {0}", Marshal.GetLastWin32Error());
                return;
            }
            Log.Write("钩子已装上（tid={0}）", _hookThreadId);

            Native.MSG msg;
            while (Native.GetMessage(out msg, IntPtr.Zero, 0, 0) > 0)
            {
                Native.TranslateMessage(ref msg);
                Native.DispatchMessage(ref msg);
            }
            Native.UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;
        }

        /// <summary>
        /// 钩子回调的唯一入口：只负责把异常挡在 native 边界之内。
        /// ★v151：托管异常一旦穿过 CallNextHookEx，CLR 会 FailFast 直接杀进程
        /// （托盘图标突然消失、日志一个字没有），已吞掉的键也不会还原。
        /// </summary>
        private IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try { return HookProcCore(nCode, wParam, lParam); }
            catch (Exception e)
            {
                try { Log.Write("钩子回调异常（已拦下，按键放行）：{0}", e); } catch { }
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
            }
        }

        private IntPtr HookProcCore(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0) return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            int msg = wParam.ToInt32();
            bool isDown = msg == Native.WM_KEYDOWN || msg == Native.WM_SYSKEYDOWN;
            bool isUp = msg == Native.WM_KEYUP || msg == Native.WM_SYSKEYUP;
            if (!isDown && !isUp) return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            var kb = (Native.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(
                lParam, typeof(Native.KBDLLHOOKSTRUCT));

            // 注入的键（我们自己发的空格/括号/回车）不算"用户按键"：否则会自己触发自己
            // （探针 --enterprobe 通过 AcceptInjectedVk 单独放行它要发的那几个键）
            if ((kb.flags & Native.LLKHF_INJECTED) != 0 &&
                AcceptInjectedVk != -1 && (int)kb.vkCode != AcceptInjectedVk)
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            int vk = (int)kb.vkCode;

            // ---- keyup：跟被吞掉的 keydown 配对的也一起吞 ----
            // ★Python 版同一个坑：只吞 keydown 的话，VBE 会收到一个没有 down 的 up
            //   （Ctrl+D / Ctrl+K 这类会被当成"半截按键"），表现为偶发的重复动作。
            if (isUp)
            {
                lock (_swallowed)
                {
                    // 只有"刚吞了 keydown 的那个键"才连带吞掉它的 keyup；过期的条目
                    // 一并清掉，绝不让它去吞不相干的 keyup（丢 keyup = 修饰键粘住）。
                    if (_swallowed.TryGetValue(vk, out long at))
                    {
                        _swallowed.Remove(vk);
                        if (Environment.TickCount64 - at <= SwallowPairMs) return (IntPtr)1;
                    }
                }
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
            }

            // ★时标打在【这里】：放到别的回调里就晚了（那时钩子已经把事件
            //   post 给监听器，文本可能已经变了）
            _lastKeyTs = Environment.TickCount64;

            bool ctrl = Native.Down(Native.VK_CONTROL);
            bool shift = Native.Down(Native.VK_SHIFT);
            bool alt = Native.Down(Native.VK_MENU);
            bool win = Native.Down(Native.VK_LWIN) || Native.Down(Native.VK_RWIN);

            // 不在 VBE 代码窗（或片段窗占着焦点）：一个键都不碰。
            // ★片段窗那道闸必须看【实时前台】而不是"窗子开着就算"：_codePaneFocused
            //   是轮询算的，片段窗刚弹出的几十上百毫秒里还是 true，那时在搜索框打字
            //   会被当成在代码里打字。GetForegroundWindow 是纯 Win32，没成本。
            // ★该管的组合键被焦点判定跳过时要记一笔 —— 那是"快捷键全不灵"最常见的真因。
            bool snipActive = SnippetFormActive();
            if (!_codePaneFocused || snipActive)
            {
                if (ctrl && IsHotkeyDiagVk(vk))
                    Log.Write("热键被跳过：焦点判定=不在 VBE 代码窗（vk=0x{0:X2} shift={1} codePane={2} snippet={3}）",
                        vk, shift, _codePaneFocused, snipActive);
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
            }

            if (win) return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            // ★v179：只记"能改空白"的键。敲 `1` 不是改空白，不能拿它挡住光标搬家 ——
            //   用户打完最后一个字马上按 Ctrl+S（间隔一两百毫秒）是最自然的操作，
            //   而 VBE 保存时正是"重排这一行 + 光标按原列号留着"。
            if (vk == Native.VK_SPACE || vk == Native.VK_TAB ||
                vk == Native.VK_BACK || vk == Native.VK_DELETE)
            {
                _lastBlankTs = Environment.TickCount64;
            }
            if (ctrl && vk == Native.VK_S) _lastSaveTs = Environment.TickCount64;

            // ---- 数"按了可能写字的键、文档却没变"的次数 ----
            // 那就是输入法正在组字的信号（拼音被吃掉、还没上屏）。组字状态下
            // 回车是【上屏键】、退格是【删拼音】，抢了就是"字没上屏反而多一行"。
            // 文档一变（轮询看到）计数就清零，所以正常打字时它总是 0。
            if (Editing.IsTextKey(vk) && !ctrl && !alt)
            {
                int n = _pendingKeys;
                if (n < 99) _pendingKeys = n + 1;
                _pendingAt = Environment.TickCount64;
            }

            // ==============================================================
            // 1. 候选窗可见时的导航键
            // ==============================================================
            if (PopupAlive && !ctrl && !alt)
            {
                if (vk == Native.VK_DOWN && !shift)
                {
                    Post(() => _eng.Move(1));
                    return Swallow(vk);
                }
                if (vk == Native.VK_UP && !shift)
                {
                    Post(() => _eng.Move(-1));
                    return Swallow(vk);
                }
                // ★Shift+↑/↓ = 跳出候选列表 + 光标上/下移一行。
                //   刻意【不】走上面"在列表里选词"那条：VBE 原生的 Shift+↑/↓ 是
                //   扩选，不是移光标，所以只能由我们用 COM 落光标（并吞掉按键）。
                if ((vk == Native.VK_UP || vk == Native.VK_DOWN) && shift && !NoShiftNav)
                {
                    int delta = vk == Native.VK_UP ? -1 : 1;
                    Post(() =>
                    {
                        _eng.Hide();
                        if (!_be.MoveCaretLine(delta)) Native.SendVk(vk, false, true);
                    });
                    return Swallow(vk);
                }
                if (vk == Native.VK_NEXT)
                {
                    Post(() => _eng.Move(_eng.PageSize));
                    return Swallow(vk);
                }
                if (vk == Native.VK_PRIOR)
                {
                    Post(() => _eng.Move(-_eng.PageSize));
                    return Swallow(vk);
                }
                // ★v147：空格 = "这个词打完了"，VBE 自己的列表也是这一下收的。
                //   轮询比的是 Rstrip 之后的文本（行尾空白刻意不算文档变化），于是
                //   `Set ` 与 `Set` 一字不差 ⇒ 走不到 Trigger，列表就一直挂着
                //   （用户报的 "set + 空格不收"）。这里只收窗、不吞键。
                if (vk == Native.VK_SPACE)
                {
                    Post(() => _eng.Hide());
                    return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }
                if (vk == Native.VK_TAB && !shift)
                {
                    Post(() => { if (!_eng.Accept()) Native.SendVk(Native.VK_TAB); });
                    return Swallow(vk);
                }
                if (vk == Native.VK_RETURN && !shift)
                {
                    // ★回车【不】确认候选 —— 回车在 VBE 里是换行，误触代价太大（确认只用 Tab），这里只收列表。
                    // ★v142：不要在这里 return。用户敲完 `Then` 立刻回车时轮询还没看到最后那个
                    //   n，候选窗正巧还挂着，return 就等于把回车整个交回 VBE（只换行、不补 End If）。
                    Post(() => _eng.Hide());
                }
                if (vk == Native.VK_ESCAPE)
                {
                    Post(() => _eng.Hide());
                    return Swallow(vk);
                }
                if (vk == Native.VK_LEFT || vk == Native.VK_RIGHT)
                {
                    // 移动光标即放弃本次补全。刻意【不】吞键 —— 光标要照常移动。
                    Post(() => _eng.Hide());
                    return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                }
            }

            // ==============================================================
            // 2. 手动唤出
            // ==============================================================
            if (ctrl && !shift && !alt && vk == Native.VK_SPACE)
            {
                Post(() =>
                {
                    var snap = _be.Snapshot();
                    if (snap != null) _eng.Trigger(true, snap);
                });
                return Swallow(vk);
            }

            // ==============================================================
            // 3. Ctrl+Q：代码片段窗（开关）
            //    排在所有"改写文本"的分支之前 —— 它是纯粹的开/关界面动作，
            //    与候选窗是否可见无关。Ctrl+Q 在 VBE 里本来什么也不做。
            // ==============================================================
            if (ctrl && !shift && !alt && vk == 'Q' && !NoSnippets)
            {
                Post(ToggleSnippets);
                return Swallow(vk);
            }

            // ==============================================================
            // 4. 会改写文本的组合键
            // ★"做成了才算吃掉了键"：命令返回 false（没快照 / COM 失败 / 没有可落的位置）
            //   就把键原样送回去，绝不"吞了但什么都没发生"。
            // ★`/` 同时认 VK_OEM_2 与小键盘 VK_DIVIDE，并允许 Shift（AZERTY 的 `/` 自带 Shift）。
            // ==============================================================
            if (ctrl && !alt && Native.IsSlashKey(vk) && !NoComment)
            {
                Post(() =>
                {
                    // ★别在失败时补发一个裸 `/`：SendVk 遇到"Ctrl 正被物理按住"
                    //   会刻意不发 Ctrl，于是补发的其实是"/"这个字符，直接往用户
                    //   代码里插一个斜杠 —— 比"什么都没发生"更糟。VBE 里 Ctrl+/
                    //   本来也不做任何事，所以这里只是记一笔。
                    if (!Commands.ToggleComment(_be))
                        Log.Write("Ctrl+/ 没做成：拿不到选区 / 整段全是空行 / 一行都没写成功（键已吞掉，VBE 里本也无此功能）");
                });
                return Swallow(vk);
            }
            // ⚠️ Shift+D / Shift+K 绝不能接管 —— 那是输入大写字母
            if (ctrl && !shift && !alt && vk == 'D' && !NoDupLine)
                return SwallowAndRun('D', true, false, () => Commands.DuplicateLine(_be),
                    replayOnFail: false);
            if (ctrl && !shift && !alt && vk == 'K' && !NoDelLine)
                return SwallowAndRun('K', true, false, () => Commands.DeleteLine(_be),
                    replayOnFail: false);
            if (ctrl && shift && !alt && vk == Native.VK_RETURN && !NoInsertAbove)
                return SwallowAndRun(Native.VK_RETURN, true, true, () => Commands.InsertLineAbove(_be),
                    replayOnFail: false);
            // ==============================================================
            // 4c. Ctrl+↑ / Ctrl+↓（不带 Shift）：过程内的两个端点
            //     ↑ = 本过程最后一个 Dim 行行尾；↓ = 收尾行上方最后一条代码行尾。
            //     （v149 已按用户要求整组删掉 Ctrl+Shift+↑/↓，这里不再有顺序约束）
            // ==============================================================
            if (ctrl && !shift && !alt && (vk == Native.VK_UP || vk == Native.VK_DOWN) && !NoCtrlNav)
            {
                int dir = vk == Native.VK_UP ? -1 : 1;
                // 判据不成立（光标在声明区 / 过程里没有 Dim / 过程体为空）就把键
                // 还回去。SendVk 在"Ctrl 正被物理按住"时不合成 Ctrl，所以实际
                // 发出的是裸 ↑/↓。
                return SwallowAndRun(vk, true, false, () => Commands.CtrlProcJump(_be, dir));
            }

            // ==============================================================
            // 5. Shift+Enter：在下方新起一行（不拆行），缩进对齐 + 补收尾
            //    ⚠️ 必须排在"回车自动缩进"之前，两条判据互斥（这条要求 Shift）
            // ==============================================================
            if (shift && !ctrl && !alt && vk == Native.VK_RETURN)
                return SwallowAndRun(Native.VK_RETURN, false, true,
                    () => _be.NewLineBelow(false, !NoAutoClose, !NoBranchAlign));

            // ==============================================================
            // 6. 回车自动缩进（缩进 / 补收尾 / 分支对齐）
            //    只有"整行注释 / 块结构开头 + 光标在行尾 + 行内无 Tab"才接管，
            //    其余一律放行给 VBE 原生回车（零风险）。
            // ==============================================================
            if (!ctrl && !shift && !alt && vk == Native.VK_RETURN && !NoEnterIndent)
            {
                string lineText;
                int caretDisp;
                bool fresh = FreshHookSnap(out lineText, out caretDisp);
                if (fresh && Editing.EnterIndentWanted(lineText, caretDisp))
                    return SwallowAndRun(Native.VK_RETURN, false, false,
                        () => _be.NewLineBelow(true, !NoAutoClose, !NoBranchAlign));

                // ★v142：快照"落后一两格"时也要接管 —— 用户敲完 `To 3` 立刻回车，
                //   那一瞬间 _pendingKeys > 0、行文本少一个字符，上面的判据必然 false。
                //   这一档不猜结论，只"先接管"，由 NewLineBelow 用真 COM 数据重判，
                //   判不过就把回车原样还回去。
                //   ⚠️ 输入法正在组字时绝不接管 —— 那一下回车是"上屏"（走 ImmGetContext，不碰 COM）。
                if (!Native.ImeComposing(Native.FocusedVbeWindow()) &&
                    Editing.LooksLikeBlockHeaderPrefix(lineText) &&
                    caretDisp + 16 >= Rules.DispCol(lineText, (lineText ?? "").Length))
                    return SwallowAndRun(Native.VK_RETURN, false, false,
                        () => _be.NewLineBelow(true, !NoAutoClose, !NoBranchAlign));
            }

            // ==============================================================
            // 7. 退格：删整行空白 / 删一对空括号
            // ==============================================================
            if (!ctrl && !shift && !alt && vk == Native.VK_BACK &&
                (!NoBsDelLine || !NoBsDelPair))
            {
                string lineText;
                int caretDisp;
                if (FreshHookSnap(out lineText, out caretDisp))
                {
                    bool blankLine = !NoBsDelLine && Editing.BlankLineDeleteWanted(lineText, caretDisp);
                    int left, right;
                    bool pair = !NoBsDelPair && !blankLine &&
                                Editing.PairDeleteSpan(lineText, caretDisp, out left, out right);
                    if (blankLine)
                        return SwallowAndRun(Native.VK_BACK, false, false, () => _be.DeleteBlankLineHere());
                    if (pair)
                        return SwallowAndRun(Native.VK_BACK, false, false, () => _be.DeletePairHere());
                }
            }

            // ==============================================================
            // 8. 自动配对括号 / 引号（含中文标点的全角）
            // ==============================================================
            if (NoAutoPair == false && !ctrl && !alt)
            {
                char half = Native.PairCharForKey(vk, shift);
                if (half != '\0')
                {
                    // ★到这里才问输入法状态（要发一次窗口消息）。放在"已经确定要
                    //   接管这个键"之后，别每次按 Shift 都去问。
                    char ch = Native.PairCharForIme(half, Native.GetForegroundWindow());
                    Post(() =>
                    {
                        _eng.Hide();
                        if (!_be.InsertPair(ch)) Native.SendUnicodeChar(ch);
                    });
                    return Swallow(vk);
                }
            }

            // 只记录"按了 Ctrl 却没命中任何分支"的、与该管的热键相关的键 ——
            // 用来排查"快捷键不生效"。日常打字（Ctrl+C / Ctrl+V 之类）不进日志。
            if (ctrl && IsHotkeyDiagVk(vk))
                Log.Write("热键未命中：vk=0x{0:X2} ctrl={1} shift={2} alt={3} popup={4}",
                    vk, ctrl, shift, alt, _popupVisible);

            return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }

        /// <summary>吃掉这一下按键（keydown），并记下 vk 好把配对的 keyup 也吞掉。</summary>
        private IntPtr Swallow(int vk)
        {
            long now = Environment.TickCount64;
            lock (_swallowed)
            {
                // 顺手清理过期条目：正常情况下 keyup 一来就删掉了，这里只是兜底，
                // 免得"吞了 keydown 却没有 keyup"的残渣越积越多。
                if (_swallowed.Count > 0)
                {
                    List<int> dead = null;
                    foreach (var kv in _swallowed)
                        if (now - kv.Value > SwallowPairMs) (dead ??= new List<int>()).Add(kv.Key);
                    if (dead != null) foreach (int k in dead) _swallowed.Remove(k);
                }
                _swallowed[vk] = now;
            }
            return (IntPtr)1;
        }

        /// <summary>
        /// "吞掉这个键 → 去 UI 线程干事 → 干不成再把键原样还回去"（全项目这条铁律的唯一落地点）。
        /// ⚠️ 字母键与回车必须 replayOnFail = false：SendVk 见 Ctrl 被按住会刻意不合成 Ctrl，
        /// 回放出去的是裸字母/裸回车，比"什么都没发生"糟得多。
        /// </summary>
        private IntPtr SwallowAndRun(int vk, bool restoreCtrl, bool restoreShift,
            Func<bool> work, bool hidePopup = true, bool replayOnFail = true)
        {
            Post(() =>
            {
                if (hidePopup) _eng.Hide();
                if (work()) return;
                if (!replayOnFail)
                {
                    Log.Write("快捷键 vk=0x{0:X2} 没做成（键已吞掉，不回放裸字符）", vk);
                    return;
                }
                Native.SendVk(vk, restoreCtrl, restoreShift);
            });
            return Swallow(vk);
        }

        /// <summary>
        /// 钩子线程用：前台窗口是不是片段窗（或它的子控件）？
        /// 只读一个缓存的 HWND + `GetForegroundWindow` + 几次 `GetParent`，
        /// 全是纯 Win32，一次 COM 都不碰，也不去碰任何 WinForms 对象。
        /// </summary>
        private bool SnippetFormActive()
        {
            IntPtr fg = _snippetHwnd;
            if (fg == IntPtr.Zero) return false;
            IntPtr h = Native.GetForegroundWindow();
            for (int i = 0; i < 8 && h != IntPtr.Zero; i++)
            {
                if (h == fg) return true;
                IntPtr p = Native.GetParent(h);
                if (p == h) break;
                h = p;
            }
            return false;
        }

        /// <summary>钩子线程用：拿"轮询留下的快照"。快照过期 / 输入法可能正在组字时一律
        /// 返回 false —— 宁可这一次不接管（原生行为无害），也绝不拿旧行文本去赌
        /// 用户正在敲的那一行。
        /// </summary>
        private bool FreshHookSnap(out string lineText, out int caretDisp)
        {
            lineText = "";
            caretDisp = 0;
            var s = _hookSnap;
            if (s == null) return false;
            // 800ms：与 Python 版的 ENTER_CTX_MAX_AGE 同口径。在代码窗里打字时
            // 轮询是 50ms 一轮，这个上限其实是"轮询卡住了 / 刚切过焦点"的兜底。
            if (Environment.TickCount64 - _hookSnapTicks > 800) return false;
            // ★v142：出参先填上（即便最后返回 false）—— 回车的"快照落后"那一档
            //   还要拿这行的文本去判"看起来像不像块头"，见 HookProc 里的说明。
            lineText = s.Text ?? "";
            caretDisp = s.CaretCol;
            // 计数超过 2 秒还挂着 ⇒ 不是"输入法在组字"，是那次变化我们没采到。
            // 不清掉的话回车缩进 / 退格接管会被永久挡死（宁可放过这一次）。
            if (_pendingKeys > 0 && Environment.TickCount64 - _pendingAt > 2000) _pendingKeys = 0;
            if (_pendingKeys > 0) return false;
            return caretDisp > 0;
        }

        /// <summary>Ctrl+Q：开 / 关片段窗（UI 线程）。开窗位置取【开窗之前】的光标。</summary>
        private void ToggleSnippets()
        {
            try
            {
                if (_snippetForm != null && _snippetForm.Visible)
                {
                    _snippetForm.Close_();
                    return;
                }
                if (_snippetForm == null)
                {
                    _snippetStore = new SnippetStore();
                    _snippetForm = new SnippetForm(_snippetStore, code => _be.InsertSnippet(code));
                    _snippetForm.OnClosedCallback = () =>
                    {
                        _snippetHwnd = IntPtr.Zero;    // 钩子线程的实时判据一并失效
                        try { _be.FocusVbe(); } catch { }
                    };
                    // 用户直接点标题栏的 × 时 FormClosing 会转成 Hide + OnClosed，
                    // 但句柄被系统回收时不会 —— 这里补一道，免得 _snippetHwnd 卡住
                    _snippetForm.FormClosed += (s, e) => _snippetHwnd = IntPtr.Zero;
                }
                _eng.Hide();

                // 位置必须在窗子显示【之前】取：一显示前台就变了，插入符矩形就没了
                int x, top, bottom;
                bool has = Native.CaretRect(out x, out top, out bottom);
                // ★句柄在这里（UI 线程上）逼出来并存给钩子线程用；窗子还没建句柄
                //   的时候 `_snippetHwnd` 就是 0，钩子那条实时判据会退化成"不拦"，
                //   而那时前台还在 VBE 代码窗 —— 正是我们想要的行为。
                _ = _snippetForm.Handle;
                _snippetHwnd = _snippetForm.Handle;
                _snippetForm.OpenAtCaret(has, x, top, bottom);
            }
            catch (Exception e)
            {
                _snippetHwnd = IntPtr.Zero;
                Log.Write("片段窗打开失败：{0}", e.Message);
            }
        }

        /// <summary>是不是"我们该管的"组合键（只用于诊断日志，免得每天打的字都进日志）。</summary>
        private static bool IsWatchedHotkey(int vk) =>
            vk == 'D' || vk == 'K' || vk == 'Q' ||
            Native.IsSlashKey(vk) ||
            vk == Native.VK_RETURN || vk == Native.VK_UP || vk == Native.VK_DOWN ||
            vk == Native.VK_SPACE;

        /// <summary>
        /// 诊断用：Ctrl + 任何一个符号键（OEM 0xBA~0xDF 与小键盘除号）没命中分支时也留痕。
        /// ★v143：以前只记 IsWatchedHotkey 里那几个 vk，"用户按的其实是另一个 vk"在日志里一个字都没有。
        /// </summary>
        private static bool IsHotkeyDiagVk(int vk) =>
            IsWatchedHotkey(vk) || (vk >= 0xBA && vk <= 0xDF);

        /// <summary>
        /// 候选窗"现在真的还能用"吗 —— 看心跳，不只看 _popupVisible。
        /// ★UI 线程卡住（COM 阻塞 / Excel 模态框 / 断点）时轮询停了而标志还是 true，
        ///   那段时间 ↑↓/Tab/Esc 全被吞、屏幕毫无反应（用户感觉键盘死了）。3 秒没心跳就算没了。
        /// </summary>
        private bool PopupAlive =>
            _popupVisible && Environment.TickCount64 - _popupBeat < 3000;

        private void Post(Action a)
        {
            try
            {
                _disp.BeginInvoke((Action)(() =>
                {
                    // 每个键盘动作也是一次"COM 使用期"：WPS 上用完必须释放
                    ComCtx.Begin();
                    try { a(); }
                    catch (Exception e) { Log.Write("UI 动作异常：{0}", e.Message); }
                    finally { ComCtx.End(); }
                }));
            }
            catch (Exception e) { Log.Write("投递动作失败：{0}", e.Message); }
        }

        // ------------------------------------------------------------------
        // 轮询
        // ------------------------------------------------------------------

        /// <summary>纯 Win32：焦点是不是在 VBE 的代码窗格里（0.02ms，不碰 COM）。</summary>
        private static bool InVbeCodePane()
        {
            IntPtr frame = Native.FindWindow(VbeFrameClass, null);
            if (frame == IntPtr.Zero) return false;
            uint pid;
            uint tid = Native.GetWindowThreadProcessId(frame, out pid);
            IntPtr focus = Native.FocusedHwndOf(tid);
            if (focus == IntPtr.Zero) return false;

            string cls = Native.ClassNameOf(focus);
            foreach (var p in VbePopupClasses)
                if (string.Equals(cls, p, StringComparison.OrdinalIgnoreCase)) return true;

            IntPtr h = focus;
            for (int i = 0; i < 12 && h != IntPtr.Zero; i++)
            {
                if (string.Equals(Native.ClassNameOf(h), CodePaneClass, StringComparison.OrdinalIgnoreCase))
                    return true;
                if (h == frame) break;
                IntPtr p = Native.GetParent(h);
                if (p == h) break;
                h = p;
            }
            // 拿不准（类名不认识）：退回"前台窗口属于 VBE 进程"
            string fgCls = Native.ClassNameOf(Native.GetForegroundWindow());
            return string.Equals(fgCls, VbeFrameClass, StringComparison.OrdinalIgnoreCase);
        }

        private void PollOnce()
        {
            if (!_running) return;

            // 最近使用记录：延迟合并写盘（2 秒最多一次）。之前是"每接受一个候选就
            // 整文件重写一次"，家目录落在 OneDrive / 网络盘上时每次 Tab 都卡一下。
            try
            {
                if (Environment.TickCount64 - _recentFlushAt > 2000)
                {
                    _recent.Flush();
                    _recentFlushAt = Environment.TickCount64;
                }
            }
            catch (Exception e) { Log.Write("recent 写盘失败：{0}", e.Message); }

            // ★WPS 上"用完即放"：整轮轮询期间允许缓存，轮询结束一定释放（见 ComCtx.End）。
            ComCtx.Begin();
            bool typingNow = false;
            try
            {
                bool focused = InVbeCodePane();
                _codePaneFocused = focused;
                // 心跳：只要 UI 线程还活着、轮询还在跑，候选窗的"可用性"就续期一次
                if (_popupVisible) _popupBeat = Environment.TickCount64;

                // 用户已经在关 WPS（主窗没了/隐藏），只剩 VBE 那个独立窗口还挂着 ——
                // 这时候一次 COM 调用都不能发，否则正好卡在它退出流程里。
                if (ComCtx.IsWpsHost && Native.WpsMainWindowGone())
                {
                    ComCtx.Release();
                    if (_popupVisible) _eng.Hide();
                    _lastSnap = null;
                    _hookSnap = null;
                    _lastModName = "";
                    SetInterval(250);
                    return;
                }

                if (!focused)
                {
                    if (_popupVisible) _eng.Hide();
                    _lastSnap = null;
                    _hookSnap = null;
                    _lastModName = "";
                    // 焦点已经不在代码窗：那些"按了键却没落进文档"的计数从此无从
                    // 求证（钩子在这之后也不再累加），留着只会白挡 2 秒的回车/
                    // 退格。Python 版同样在"刚切回 VBE"时清零。
                    _pendingKeys = 0;
                    SetInterval(250);
                    return;
                }

                // ★v180：刚按过键就别吃快照缓存 —— 这一行可能已经被宿主改过了，而改法
                //   恰好叫缓存命中。VBE 保存（Ctrl+S）时会把当前行重排空白，**却把光标
                //   留在原列号上**；缓存的键就是"行号 + 光标列"，于是新行最多晚
                //   SNAP_MAX_AGE（默认 1 秒）才看得见。用户在这一秒里接着打字：
                //   那半截词被当成真词弹窗，光标也没人搬 —— 就是用户报的"又复现"。
                //   代价：按键后这段窗口里每轮多读一次行文本（缓存本来也命不中，光标在动）。
                if (Environment.TickCount64 - _lastKeyTs <= KeyTriggerWindowMs)
                    _be.InvalidateSnapshot();

                Snapshot snap = _be.Snapshot();
                if (snap == null)
                {
                    if (_popupVisible) _eng.Hide();
                    _hookSnap = null;
                    ExplainVbeDenied();
                    SetInterval(150);
                    return;
                }
                _deniedShown = false;

                // ★v139：VBE 自己的提示词列表一出现，我们立刻收 —— 不能只靠
                //   "下次触发时让位"。触发那一刻 VBE 的列表往往还没画出来（它等
                //   WM_TIMER 才 Show），两边会并排挂在屏幕上几百毫秒，用户看到的
                //   就是一个压着另一个。这里每轮（打字时 50ms）查一次，
                //   缓存 150ms，开销约 0。
                if (_popupVisible && Native.VbeMemberListVisible())
                {
                    _eng.Hide();
                    // 日志限速：每 2 秒最多一条，别把 log 刷满
                    long now = Environment.TickCount64;
                    if (now - _lastYieldLog > 2000)
                    {
                        _lastYieldLog = now;
                        Log.Write("让位：VBE 的成员列表已出现，收起我们的候选窗");
                    }
                }

                typingNow = Environment.TickCount64 - _lastKeyTs < 900;
                // ★WPS 上单轮成本比 Office 高，间隔放宽一档：空闲 300ms / 打字 80ms。
                SetInterval(typingNow
                    ? (ComCtx.IsWpsHost ? 80 : 50)
                    : (ComCtx.IsWpsHost ? 300 : 150));

                // 给钩子线程留一份"最近的行 + 光标列"：回车缩进 / 退格那几条判据
                // 全靠它（钩子线程一次 COM 都不许碰）。
                _hookSnapTicks = Environment.TickCount64;
                _hookSnap = snap;

                bool changed = _lastSnap == null
                    || snap.Line != _lastSnap.Line
                    || !string.Equals(Rules.Rstrip(snap.Text), Rules.Rstrip(_lastSnap.Text), StringComparison.Ordinal)
                    || !string.Equals(snap.Module, _lastModName, StringComparison.OrdinalIgnoreCase);

                if (changed)
                {
                    // 文档变了 ⇒ "按了键却没落进文档"的计数清零（见 _pendingKeys 的说明）
                    _pendingKeys = 0;
                    // 上一轮的快照（下面要拿它的文本/光标列算"这一行被重排了没有"）
                    Snapshot prev = _lastSnap;

                    // 换模块 = 不是打字，只是切了代码窗，不该弹
                    bool modSwitched = !string.Equals(snap.Module, _lastModName, StringComparison.OrdinalIgnoreCase)
                                       && _lastSnap != null;

                    // ★v141：行号变了 ⇒ 光标是被"搬"过去的，不是在打字。
                    //   鼠标点另一行的变量时"行号 + 行文本"全变，和敲一个字长得一样，
                    //   "最近 1.2 秒按过键"拦不住（用户常打完字顺手点一下变量）。
                    //   真打字时行号绝不会变（回车换到的是空行，本来也没词可提示）。
                    bool lineJumped = _lastSnap != null && snap.Line != _lastSnap.Line;
                    int prevLine = _lastSnap == null ? 0 : _lastSnap.Line;

                    _lastSnap = snap;
                    _lastModName = snap.Module;

                    // ★宿主把这一行重排了空白：Ctrl+S 会顺手格式化当前行（`n=n+1` → `    n = n + 1`），
                    //   而光标只按原列号保留，于是跑到第二个 `n` 后面。两件事都要收：
                    //     ① 别弹 —— 那屏候选是对着一个不存在的"正在输入的词"算出来的；
                    //     ② 光标按"非空白字符序号"搬回等价位置，否则接着打字就插错地方。
                    //   ★与"人手敲空格"的区分（v179 重写）：
                    //     ① 刚按过 Ctrl+S —— VBE 保存时一定会格式化当前行，最硬的前兆；
                    //     ② 空白个数变了、且最近没按过空白键（空格/Tab/退格/Delete）。
                    //        人一次只能动一个空格，一次轮询出现多寡变化只能是宿主干的。
                    //     ★v174 那版用的是"最近 400ms 没有任何编辑键"，于是"敲完 `1`
                    //       立刻 Ctrl+S"整条被跳过 —— 恰好把用户最常做的动作排除在外。
                    if (!NoRespacedFix && prev != null && !modSwitched && !lineJumped &&
                        Rules.SameCodeIgnoringSpace(snap.Text, prev.Text))
                    {
                        if (_popupVisible) _eng.Hide();
                        long nowTs = Environment.TickCount64;
                        int dSpace = Rules.SpaceCount(snap.Text) - Rules.SpaceCount(prev.Text);
                        bool saveRecent = nowTs - _lastSaveTs <= SaveWindowMs;
                        bool blankKeyRecent = nowTs - _lastBlankTs <= NoBlankKeyMs;
                        bool hostRespaced = saveRecent || (dSpace != 0 && !blankKeyRecent);
                        if (hostRespaced)
                        {
                            int oldCaret = Rules.CharColFromDisp(prev.Text ?? "", prev.CaretCol);
                            int want = Rules.RemapCaretOverSpaces(prev.Text ?? "", oldCaret, snap.Text ?? "");
                            int now = Rules.CharColFromDisp(snap.Text ?? "", snap.CaretCol);
                            if (want > 0 && want != now && _be.SetCaret(snap.Line, want))
                            {
                                Log.Write("宿主重排：L{0} 光标 {1}->{2}（行内空白被重排，不弹）",
                                    snap.Line, now, want);
                                var fixedSnap = _be.Snapshot();
                                if (fixedSnap != null)
                                {
                                    _lastSnap = fixedSnap;
                                    _hookSnap = fixedSnap;
                                }
                            }
                        }
                        return;
                    }

                    if (!modSwitched)
                    {
                        _eng.InvalidatePool();
                        // ★文本变了不等于"用户在打字"：鼠标右键粘贴、拖拽一段
                        //   代码进来、别的程序通过 COM 改了模块，都会走到这里。
                        //   那些场景一律只收起、不弹（Python 版 REQUIRE_KEY_TRIGGER）。
                        //   判据是"最近 1.2 秒内按过键"，以及"输入法可能正在组字"。
                        bool keyTriggered = !RequireKeyTrigger
                                            || _pendingKeys > 0
                                            || Environment.TickCount64 - _lastKeyTs <= KeyTriggerWindowMs;
                        if (!Rules.AutoTriggerWanted(true, modSwitched, keyTriggered, lineJumped))
                        {
                            if (_popupVisible) _eng.Hide();
                            // ★只有"本来会弹、因为跳行才没弹"这一种记一笔（限速 2 秒）。
                            //   用户再报"点了变量弹窗"时，日志里能直接看出是漏判还是别的原因。
                            if (lineJumped && keyTriggered)
                            {
                                long now = Environment.TickCount64;
                                if (now - _lastJumpLog > 2000)
                                {
                                    _lastJumpLog = now;
                                    Log.Write("不弹：光标跳行 L{0}->L{1}（鼠标单击 / 方向键），不是打字",
                                        prevLine, snap.Line);
                                }
                            }
                        }
                        else
                        {
                            // ★不要在这里先 Hide 一下再 Trigger：Hide 会真的把窗口
                            //   收起来，Trigger 再 Show —— 用户每敲一个字都看到候选窗
                            //   "闪一下"。Trigger 内部该收的时候自己会收（打全即收 /
                            //   空结果 / 让位）。
                            _eng.Trigger(false, snap);
                        }
                    }
                }
                else if (_popupVisible)
                {
                    // 文本没变、光标动了 = 移动光标，不是打字。
                    // ★v159：以前这里无条件 ShowPopupAtCaret()，于是同行里点一下鼠标 /
                    //   按左右键 / 按 Home，候选窗就跟着光标跑、列表一直挂着。一律收起。
                    if (_lastSnap == null || _lastSnap.CaretCol != snap.CaretCol)
                    {
                        if (_popupVisible)
                            Log.Write("收起候选窗：光标在行内移动（列 {0}->{1}），不是打字",
                                _lastSnap == null ? -1 : _lastSnap.CaretCol, snap.CaretCol);
                        _eng.Hide();
                    }
                    _lastSnap = snap;      // 让比较基准前进，别每轮都判成"动了"
                }
            }
            catch (Exception e)
            {
                Log.Write("轮询异常：{0}", e.Message);
            }
            finally
            {
                ComCtx.End(typingNow);
            }
        }

        private void SetInterval(int ms)
        {
            if (_timer == null) return;
            if (_timer.Interval != ms) _timer.Interval = ms;
        }

        /// <summary>
        /// 焦点在代码窗、却拿不到任何快照，而宿主有实例、VBE 窗口也在 —— 只可能是 AccessVBOM
        /// 没开（WPS 默认如此）。★先自动打开，但要提示"重开 WPS"（它只在启动时读这个值）；
        /// ★用户在托盘里自己关掉的不许动（VbaTrust.UserDisabled）。
        /// </summary>
        private void ExplainVbeDenied()
        {
            if (!ComCtx.VbeAccessDenied || _deniedShown) return;
            _deniedShown = true;

            if (VbaTrust.UserDisabled)
            {
                Log.Write("拿不到 VBE：AccessVBOM 是用户自己关掉的（{0}），不自动开启", "trust-off.txt");
                ShowBalloon("VBA 工程访问是你关掉的",
                    "托盘菜单「允许访问 VBA 工程」现在开着还是关着，你自己决定 —— 关着时补全在 WPS / Excel 里不工作。\r\n"
                    + "要恢复：右键托盘图标点一次那一项，然后重开宿主。");
                return;
            }

            int changed = 0;
            if (VbaTrust.IsBlocked()) changed = VbaTrust.EnableAll();
            Log.Write("拿不到 VBE：AccessVBOM 被禁，自动开启 {0} 处（当前 {1}）",
                changed, VbaTrust.Describe());

            ShowBalloon("WPS：需要允许访问 VBA 工程",
                changed > 0
                    ? "已自动开启「信任对 VBA 工程对象模型的访问」。\r\n请保存并重新打开 WPS（补全才能生效）。"
                    : "请右键托盘图标 →「允许访问 VBA 工程（WPS / Office）」，然后重开 WPS。");
        }

        private void ShowPopupAtCaret()
        {
            int x, top, bottom;
            if (!Native.CaretRect(out x, out top, out bottom))
            {
                // 拿不到光标就收起 —— 绝不猜一个位置糊上去
                _eng.Hide();
                return;
            }
            // ★ShowAt 是"先定位、再显示"：分开写（Show 之后才 PlaceAt）会让窗口
            //   在屏幕左上角闪一下再跳过来。
            _popup.ShowAt(x, top, bottom);
        }

        // ------------------------------------------------------------------
        // 退出
        // ------------------------------------------------------------------

        /// <summary>
        /// 退出。★幂等：托盘"退出"和 `Application.Run` 返回后都会调一次，
        /// 第二次再对已 Dispose 的窗子 Hide/Dispose 只会制造 ObjectDisposedException。
        /// </summary>
        public void Shutdown()
        {
            if (Interlocked.Exchange(ref _shutdownDone, 1) != 0) return;
            _running = false;
            try { _timer?.Stop(); _timer?.Dispose(); _timer = null; } catch { }
            try { _recent.Flush(); } catch { }
            try
            {
                if (_hookThreadId != 0)
                    Native.PostThreadMessage(_hookThreadId, Native.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            }
            catch { }
            try { if (_hookHandle != IntPtr.Zero) Native.UnhookWindowsHookEx(_hookHandle); } catch { }
            try { _showEvent?.Dispose(); } catch { }
            try { _popup.Hide(); _popup.Dispose(); } catch { }
            try { _snippetForm?.Hide(); _snippetForm?.Dispose(); } catch { }
            try { _snippetStore?.Dispose(); } catch { }
            try { _settingsForm?.Hide(); _settingsForm?.Dispose(); } catch { }
            try
            {
                if (_tray != null) { _tray.Visible = false; _tray.Dispose(); _tray = null; }
            }
            catch { }
            Log.Write("已退出");
            Log.Flush();          // 队列里最后几行别丢（后台线程是后台线程，进程一退就没了）
            ExitThread();
        }
    }
}
