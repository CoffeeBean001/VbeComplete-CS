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
    /// <summary>真机声明探针：候选位置判据、Ctrl+S 重排后光标。</summary>
    internal static partial class Program
    {
        /// <summary>
        /// 真机探针：在 VBE 里逐字符敲"声明位置 + 第二个关键字"的各种组合，看候选窗弹不弹。
        /// ★必须上真机：selftest / enginetest 只看文本，而用户走的是 钩子 → 快照 → Trigger ——
        ///   快照落后一格、自动空格、输入法组字，任一环节偏一下结论就不同。
        /// ★脚手架用完必删（`Sub VbcEnter`，见 CleanupScratch）。
        /// </summary>
        private static int DeclProbe(string only = null)
        {
            int onlyNo = 0;
            if (!string.IsNullOrEmpty(only)) int.TryParse(only, out onlyNo);
            var sb = new StringBuilder();
            var gate = new object();
            bool allOk = true;
            Action<string, bool, string> check = (name, ok, detail) =>
            {
                if (!ok) allOk = false;
                lock (gate) sb.AppendLine((ok ? "  PASS  " : "  FAIL  ") + name + "   " + detail);
            };

            sb.AppendLine("=== decl probe（真机 / 逐字符敲声明位置） ===");
            sb.AppendLine("（每个用例：摆好半截行 -> 逐字符敲剩下的 -> 看候选窗弹没弹）");

            var app = new App();
            var be = new VbeBackend();
            int leftover = CleanupScratch(be);
            string orig = be.ModuleText();
            int baseCount = be.LineCount();
            sb.AppendLine("module lines = " + baseCount + "  len = " + (orig == null ? -1 : orig.Length) +
                          "  上次残留块清理 = " + leftover);
            if (baseCount < 0 || orig == null)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "declprobe.log"),
                    sb.ToString(), new UTF8Encoding(false));
                return 1;
            }

            int at = baseCount + 1;
            RebuildScaffold(app, be, at, "    x = 1");
            be.FocusVbe();
            app.Start();

            // 说明, 半截行（光标停在行尾）, 接着敲的字, 期望弹?, 必须有, 不许有
            // ★空格必须算在"接着敲的字"里：VBE 把行尾空格抹掉（`    ReDim Preserve ` 存回去
            //   就是 `    ReDim Preserve`），跟着敲的字会**粘**在关键字后面 ——
            //   那测的是"一个长标识符"，不是"关键字 + 变量名"。
            var cases = new (string name, string header, string typed, bool want, string[] must, string[] bad)[]
            {
                ("ReDim 后补 Preserve", "    ReDim P", "re", true, new[] { "Preserve" }, null),
                // ★v176：用户报"写 for i，这里的 i 为什么不提示"。单字符的变量名打完时，
                //   `For` 后面的位置已经把内建全剔了，候选只剩它自己一条 —— 旧版被
                //   "打全即收"（HIDE_ON_COMPLETE）当成"名字已经敲完"立刻收窗，看起来
                //   就是"没提示"。单字符不算打全（见 Engine 里 _wordComplete 那条）。
                ("For 后敲单字符变量（i）", "    For", " i", true, new[] { "i" }, null),
                // ★v170：`ReDim` 后面命名的是**已经存在**的数组 ⇒ 该弹名字（用户报的就是这里）。
                //   `bad` 里的 Resume / Reset 是旧口径的噪音：那里本该只有名字。
                ("Preserve 后敲变量名（res）", "    ReDim Preserve", " res", true,
                    new[] { "resProbe" }, new[] { "Resume", "Reset" }),
                // 打全了整个名字就收窗（"打全即收"，本来就这么设计的），不是不弹
                ("Preserve 后打全即收（resProbe）", "    ReDim Preserve", " resProbe", false, null, null),
                ("Preserve 后括号里敲变量（0 To n）", "    ReDim Preserve resProbe(0 To", " nP", true,
                    new[] { "nProbe" }, null),
                ("ReDim 后敲变量名（不带 Preserve）", "    ReDim", " res", true, new[] { "resProbe" }, null),
                // 同一类过度抑制：`Dim arr(0 To n|` 的括号里是表达式
                ("Dim 括号里敲上界（0 To n）", "    Dim arrProbe2(0 To", " nP", true,
                    new[] { "nProbe" }, null),
                // 但 `Dim` 自己的名字位照旧不弹（别修过头）
                ("Dim 名字位不弹", "    Dim", " newName", false, null, null),
                ("Dim 名字位带 As 不弹", "    Dim newName As Str", "ing", false, null, null),
                ("Private 后补 Function", "    Private F", "un", true, new[] { "Function" }, null),
                ("Private 后敲变量名", "    Private", " gCount", false, null, null),
                ("Property 后补 Get", "    Property G", "e", true, new[] { "Get" }, null),
                // ★v174：`Declare` / `Option` 只能写在**模块级** —— 过程体里本来就不该提示
                //   （模块级的那两条见下面的 modCases，那才是用户真正写这两句的位置）。
                ("过程体里不给 Declare PtrSafe", "    Declare P", "tr", false, null,
                    new[] { "PtrSafe" }),
                ("Property Get 后敲过程名", "    Public Property Get", " getT", false, null, null),
                ("Private Sub 后敲过程名", "    Private Sub", " DoIt", false, null, null),
                ("Declare PtrSafe Function 后敲过程名", "    Private Declare PtrSafe Function", " getT", false, null, null),
                ("Private WithEvents 后敲变量名", "    Private WithEvents", " clsX", false, null, null),
                ("Private Enum 后敲枚举名", "    Private Enum", " MyEnum", false, null, null),
                ("Private Type 后敲类型名", "    Private Type", " tRec", false, null, null),
                // 声明关键字不在语句开头（单行 If / 冒号后面）—— 用户实际写紧凑代码就这么写
                ("单行 If 里的 ReDim Preserve", "    If n > 0 Then ReDim Preserve", " res", true,
                    new[] { "resProbe" }, null),
                ("单行 If 里的 ReDim", "    If n > 0 Then ReDim", " res", true, new[] { "resProbe" }, null),
                ("冒号后面的 ReDim Preserve", "    x = 1: ReDim Preserve", " res", true,
                    new[] { "resProbe" }, null),
                ("单行 If 里 Then 后补 Preserve", "    If n > 0 Then ReDim P", "re", true, new[] { "Preserve" }, null),
                // v169：同一类"关键字后面只可能跟关键字"的位置
                ("End 后补 Sub", "    End S", "u", true, new[] { "Sub" }, null),
                ("Exit 后补 For", "    Exit F", "o", true, new[] { "For" }, null),
                ("On Error 后补 GoTo", "    On Error G", "o", true, new[] { "GoTo" }, null),
                ("过程体里不给 Option Explicit", "    Option E", "x", false, null,
                    new[] { "Explicit" }),
                ("Do 后补 While", "    Do W", "h", true, new[] { "While" }, null),
                ("Loop 后补 Until", "    Loop U", "n", true, new[] { "Until" }, null),
                ("For 后补 Each", "    For E", "a", true, new[] { "Each" }, null),
                ("Case 后补 Else", "    Case El", "s", true, new[] { "Else" }, null),
                // v171：位置判据不外溢 —— 行首 `Preserve` 不是 `ReDim` 的续行，
                // 这里的 `re` 该给 `ReDim`（用户报的就是这一格）
                // ★v173：用户第二张截图 —— 给了 ReDim 之后，同前缀的 `Replace` / `Reset`
                //   / `Resume` 和变量 `result` 还挂在那儿。这个位置是纯语法位，
                //   除了 `ReDim` 什么都不合法，所以 `resProbe` 一起剔。
                ("行首 Preserve 后只给 ReDim", "    Preserve r", "e", true,
                    new[] { "ReDim" },
                    new[] { "resProbe", "Replace", "Rem", "Reset", "Resume", "Return" }),
                // `Case` 后面是表达式，内建函数不许被当成"位置不对的关键字"剔掉
                ("Case 后补内建（Chr）", "    Case C", "h", true, new[] { "Chr" }, null),
                // `Declare PtrSafe Fu|` 是第三个关键字位置
                ("PtrSafe 后补 Function", "    Private Declare PtrSafe F", "u", true,
                    new[] { "Function" }, null),
                // ★v174：用户报的原话 —— "在函数里打 private，会提示"。过程体里 `Private`
                //   一定是语法错误；同前缀的 `Print` 是内建函数，照样要给（别修过头）。
                ("函数体里不给 Private、但给 Print", "    Pri", "n", true,
                    new[] { "Print" }, new[] { "Private" }),
            };

            // ★v174：只能写在**模块级**的关键字，用例必须摆在模块级（见 RebuildScaffoldTopLevel）。
            var modCases = new (string name, string line, string typed, bool want, string[] must, string[] bad)[]
            {
                // 用户报的第二格：标准模块里打 `Prop` 不该冒出 `Property`
                ("标准模块里没有 Property", "    Prop", "e", false, null, new[] { "Property" }),
                ("模块级 Declare 后补 PtrSafe", "    Declare P", "tr", true, new[] { "PtrSafe" }, null),
                ("模块级 Option 后补 Explicit", "    Option E", "x", true, new[] { "Explicit" }, null),
            };

            int imeModeAtStart = -1;
            var driver = new Thread(() =>
            {
                try
                {
                    bool focused = false;
                    for (int i = 0; i < 40 && !focused; i++)
                    {
                        app.RunOnUi(() =>
                        {
                            Native.ForceForeground(Native.FindWindow("wndclass_desked_gsk", null));
                            be.SetCaret(at + 1, 9);
                            be.InvalidateSnapshot();
                        });
                        Thread.Sleep(250);
                        focused = app.CodePaneFocused;
                    }
                    check("焦点在代码窗", focused, "CodePaneFocused=" + focused);
                    if (!focused) return;

                    int imeMode0 = Native.ImeConversionMode(Native.FocusedVbeWindow());
                    imeModeAtStart = imeMode0;
                    if (!Native.ImeAlphanumeric(Native.FocusedVbeWindow()))
                    {
                        Native.ImeSetAlphanumeric(Native.FocusedVbeWindow());
                        Thread.Sleep(200);
                    }
                    bool alnum = Native.ImeAlphanumeric(Native.FocusedVbeWindow());
                    lock (gate)
                        sb.AppendLine("IME 英数模式 = " + alnum +
                            "（conv=0x" + imeMode0.ToString("X2") + "，跑完恢复）" +
                            (alnum ? "" : "   ⚠️ 切不过去：逐字符打字会被输入法吃掉"));

                    // 采样 + 断言（脚手架里的用例和模块级那批共用这一段）
                    Action<string, bool, string[], string[], bool, string> eval =
                        (name, want, must, bad, typed, how) =>
                        {
                            // ★v172：不能"固定睡 400ms 再看一眼"。宿主的轮询是 80~300ms 一跳
                            //   （Excel 打字时快、空闲时慢），打最后一个字符到"那一轮 Trigger
                            //   跑完并画出来"可能差好几百毫秒 —— 第 12 条（Declare PtrSafe）
                            //   和第 21 条（冒号后 ReDim Preserve）都因此偶发假红过：单跑
                            //   五次全绿、全量一跑随机红一条。改成"弹到即停、超时才判不弹"。
                            bool shown = false;
                            string names = "";
                            int waited = 0;
                            while (waited < 1500)
                            {
                                Thread.Sleep(60);
                                waited += 60;
                                if (app.PopupVisible) { shown = true; names = app.PopupNames; break; }
                                // 期望不弹的用例（打全即收 / 名字位）：收窗是立即的，等
                                // 600ms 还没有动静就结案，别把整套拖成一分多钟。
                                if (!want && waited >= 600) break;
                            }
                            string[] got = ReadLines(app, be, at, 2);
                            string line = got.Length > 1 ? got[1] : "";

                            bool pass = typed && shown == want;
                            foreach (var wantName in must ?? Array.Empty<string>())
                                if (!names.Split(',').Any(x => string.Equals(x, wantName,
                                    StringComparison.OrdinalIgnoreCase)))
                                    pass = false;
                            // ★v173：`bad` 之前只是写在用例表里，**从来没被断言过** ——
                            //   于是"不该出现的候选"一直没人管（`Preserve re|` 冒出
                            //   Reset/Resume 也只能靠人眼看截图发现）。
                            foreach (var banned in bad ?? Array.Empty<string>())
                                if (names.Split(',').Any(x => string.Equals(x, banned,
                                    StringComparison.OrdinalIgnoreCase)))
                                    pass = false;
                            check(name, pass,
                                "弹=" + shown + " 候选=[" + names + "] 行=[" + line + "] via=" + how);
                        };

                    for (int k = 0; k < cases.Length; k++)
                    {
                        var c = cases[k];
                        if (onlyNo > 0 && k + 1 != onlyNo) continue;
                        // 额外塞一行 `Dim`：`ReDim Preserve resProbe` 的 resProbe 必须是
                        // **已经在作用域里**的变量，否则池子里没有它，弹不弹都验不出来。
                        RebuildScaffold(app, be, at, c.header,
                            // ★v176：`i&` 是给"单字符变量"那条用例摆的（用户报 `for i` 不提示）。
                            "    Dim resProbe() As Long, nProbe&, i&");
                        Thread.Sleep(350);
                        bool typed = TypeTyping(app, c.typed, 55, out string how);
                        eval((k + 1) + " " + c.name, c.want, c.must, c.bad, typed, how);
                    }

                    // ★v174：**模块级**那批 —— `Declare PtrSafe` / `Option Explicit` /
                    //   `Property` 这些"只能写在模块级"的位置，包进 `Sub` 就变成"过程体内"，
                    //   会被"关键字只能在模块级"那条判据正确过滤掉，测的就不是这件事了。
                    foreach (var c in modCases)
                    {
                        RebuildScaffoldTopLevel(app, be, at, c.line);
                        Thread.Sleep(350);
                        bool typed = TypeTyping(app, c.typed, 55, out string how);
                        eval("M " + c.name, c.want, c.must, c.bad, typed, how);
                    }
                }
                catch (Exception e)
                {
                    lock (gate) sb.AppendLine("EXCEPTION " + e);
                    allOk = false;
                }
                finally
                {
                    Native.ImeRestore(Native.FocusedVbeWindow(), imeModeAtStart);
                    app.RequestExit();
                }
            })
            { IsBackground = true, Name = "declprobe" };

            driver.Start();
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
                    be.ForceReloadModule();
                    // ★v174：先把 `at` 之后的尾巴清干净 —— 模块级用例写在 `Sub` 外面，
                    //   CleanupScratch 只认 `Sub VbcEnter() … End Sub`，清不掉它。
                    int tailCnt = be.LineCount();
                    if (tailCnt >= at) be.DeleteLines(at, tailCnt - at + 1);
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
                        File.WriteAllText(Path.Combine(Settings.Home, "declprobe.log"),
                            sb.ToString(), new UTF8Encoding(false));
                    }
                }
                catch (Exception e)
                {
                    lock (gate)
                    {
                        sb.AppendLine("cleanup 异常：" + e.Message);
                        sb.AppendLine("RESULT: FAIL");
                        try
                        {
                            File.WriteAllText(Path.Combine(Settings.Home, "declprobe.log"),
                                sb.ToString(), new UTF8Encoding(false));
                        }
                        catch { }
                    }
                    allOk = false;
                }
            }
            return allOk ? 0 : 1;
        }

        /// <summary>
        /// 真机探针：**宿主把一行重排了空白之后，光标该落在哪儿**。
        /// 敲完 `n=n+1` 按 Ctrl+S，VBE 会把行格式化成 `    n = n + 1`，而光标只按原列号
        /// 保留 ⇒ 跑到第二个 `n` 后面，接着打字插错地方还会弹一屏假候选。
        /// ★不真的按 Ctrl+S（会把测试行写进用户的工作簿）：直接把这一行换成格式化后的样子、
        ///   把光标摆回原列号，再让 App 自己跑轮询读快照 —— 判据在 PollOnce 里。
        /// </summary>
        private static int FmtProbe()
        {
            var sb = new StringBuilder();
            var gate = new object();
            bool allOk = true;
            Action<string, bool, string> check = (name, ok, detail) =>
            {
                if (!ok) allOk = false;
                lock (gate) sb.AppendLine((ok ? "  PASS  " : "  FAIL  ") + name + "   " + detail);
            };

            sb.AppendLine("=== fmt probe（真机 / 宿主重排空白后的光标落点） ===");

            var app = new App();
            var be = new VbeBackend();
            int leftover = CleanupScratch(be);
            string orig = be.ModuleText();
            int baseCount = be.LineCount();
            sb.AppendLine("module lines = " + baseCount + "  len = " + (orig == null ? -1 : orig.Length) +
                          "  上次残留块清理 = " + leftover);
            if (baseCount < 0 || orig == null)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "fmtprobe.log"),
                    sb.ToString(), new UTF8Encoding(false));
                return 1;
            }

            int at = baseCount + 1;
            RebuildScaffold(app, be, at, "    n = n");
            be.FocusVbe();
            app.Start();

            // 用例（就是用户截图那一格）：手写 `    n=n+1`，Ctrl+S 之后 VBE 把它格式化成
            //   `    n = n + 1`，而光标按原列号留在了第 10 列 —— 正好是第二个 `n` 后面。
            //   期望：不弹候选（那是假词），光标搬回行尾（= `1` 后面，第 14 列）。
            //   ★`=` 两侧也必须没有空格：只差 `+1` 那一段的话，光标落在 `+` 后面、取词
            //     是空的，老代码本来就不会弹 —— 那样测不到"假词弹窗"这一半。
            const string header = "    n";
            const string typed = "=n+1";
            const string formatted = "    n = n + 1";

            var driver = new Thread(() =>
            {
                int imeModeAtStart = -1;
                try
                {
                    bool focused = false;
                    for (int i = 0; i < 40 && !focused; i++)
                    {
                        app.RunOnUi(() =>
                        {
                            Native.ForceForeground(Native.FindWindow("wndclass_desked_gsk", null));
                            be.SetCaret(at + 1, header.Length + 1);
                            be.InvalidateSnapshot();
                        });
                        Thread.Sleep(250);
                        focused = app.CodePaneFocused;
                    }
                    check("焦点在代码窗", focused, "CodePaneFocused=" + focused);
                    if (!focused) return;

                    int imeMode0 = Native.ImeConversionMode(Native.FocusedVbeWindow());
                    imeModeAtStart = imeMode0;
                    if (!Native.ImeAlphanumeric(Native.FocusedVbeWindow()))
                    {
                        Native.ImeSetAlphanumeric(Native.FocusedVbeWindow());
                        Thread.Sleep(200);
                    }
                    bool alnum = Native.ImeAlphanumeric(Native.FocusedVbeWindow());
                    lock (gate)
                        sb.AppendLine("IME 英数模式 = " + alnum +
                            "（conv=0x" + imeMode0.ToString("X2") + "，跑完恢复）" +
                            (alnum ? "" : "   ⚠️ 切不过去：逐字符打字会被输入法吃掉"));

                    RebuildScaffold(app, be, at, header, "    Dim n As Long");
                    Thread.Sleep(350);
                    bool ok1 = TypeTyping(app, typed, 55, out string how);

                    // ★500ms 是算出来的，不是随手写的：
                    //   >400ms —— 让钩子那边"最近按过的编辑键"过期，否则会被当成"用户在敲空格"；
                    //   <1200ms —— 仍在 KeyTriggerWindowMs 里，也就是**旧代码在这个位置一定会弹**
                    //   （用户截图里那屏 `n / Next / New / Not …` 就是这么来的）。
                    Thread.Sleep(500);

                    var s0 = app.HookSnap;
                    string rawText = s0?.Text ?? "";
                    int rawCol = s0?.CaretCol ?? 0;
                    int rawCharCol = Rules.CharColFromDisp(rawText, rawCol);
                    lock (gate)
                        sb.AppendLine("  敲完 -> 行=[" + rawText + "] 显示列=" + rawCol +
                                      " 字符列=" + rawCharCol + " via=" + how);
                    bool typOk = ok1 && string.Equals(Rules.Rstrip(rawText),
                        Rules.Rstrip(header + typed), StringComparison.Ordinal);

                    // 摆成"保存之后"的样子：文本 = 格式化后的，光标 = 原来的列号
                    // （真机上 Ctrl+S 干的就是这两件事：重排 + 光标按原列号留着）
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        be.ReplaceLineRaw(at + 1, formatted);
                        be.SetCaret(at + 1, Math.Max(1, rawCharCol));
                        be.InvalidateSnapshot();
                    });
                    // 再"按一下无关的键" —— 真机上这一下就是 Ctrl+S 本身。
                    // ★为什么非按不可：App 的快照有 800ms 保鲜期，而保存之后**光标列没变**
                    //   （正是这个 bug），缓存整条命中 ⇒ 要等它过期才看得见新文本。这期间
                    //   "最近 1.2 秒按过键"早就过期了，弹窗那一半根本测不出来（试过一次，
                    //   A/B 关掉修复也只看到光标错位、看不到弹窗）。
                    SendKeys(app, Native.VK_SHIFT);
                    Thread.Sleep(900);

                    Snapshot s1 = null;
                    app.RunOnUi(() =>
                    {
                        be.ForceReloadModule();
                        s1 = be.Snapshot();
                    });
                    string text = s1?.Text ?? "";
                    int caret = s1 == null ? -1 : Rules.CharColFromDisp(text, s1.CaretCol);
                    bool shown = app.PopupVisible;
                    string names = app.PopupNames;
                    int want = text.Length + 1;      // `1` 后面 = 行尾
                    bool pass = typOk && !shown && s1 != null && caret == want &&
                                string.Equals(Rules.Rstrip(text), Rules.Rstrip(formatted),
                                    StringComparison.Ordinal);
                    check("保存重排后：不弹窗 + 光标回到行尾", pass,
                        "弹=" + shown + (names.Length > 0 ? "[" + names + "]" : "") +
                        " 行=[" + text + "] 光标字符列=" + caret + "（期望 " + want + "）" +
                        " 敲入=" + typOk);

                    if (!string.Equals(Rules.Rstrip(text), Rules.Rstrip(formatted),
                            StringComparison.Ordinal))
                        lock (gate) sb.AppendLine("  ⚠️ 行文本与期望不符，本条的结论仅供参考");
                }
                catch (Exception e)
                {
                    lock (gate) sb.AppendLine("EXCEPTION " + e);
                    allOk = false;
                }
                finally
                {
                    Native.ImeRestore(Native.FocusedVbeWindow(), imeModeAtStart);
                    app.RequestExit();
                }
            })
            { IsBackground = true, Name = "fmtprobe" };

            driver.Start();
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
                        File.WriteAllText(Path.Combine(Settings.Home, "fmtprobe.log"),
                            sb.ToString(), new UTF8Encoding(false));
                    }
                }
                catch (Exception e)
                {
                    lock (gate)
                    {
                        sb.AppendLine("cleanup 异常：" + e.Message);
                        sb.AppendLine("RESULT: FAIL");
                        try
                        {
                            File.WriteAllText(Path.Combine(Settings.Home, "fmtprobe.log"),
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
