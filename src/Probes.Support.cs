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
    /// <summary>真机探针的公用脚手架：改临时行、灌键、读回。</summary>
    internal static partial class Program
    {
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
                // ★上限别卡太死：用例 G 的块有 42 行，原来写的是 head+40 —— 于是
                //   找不到 End Sub、整个块留在用户模块里（cleanup 0 / restored False）。
                for (int i = head + 1; i <= Math.Min(cnt, head + 500); i++)
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

        /// <summary>
        /// 把用例摆在**模块级**：`用例行` + 一个空的 `Sub VbcEnter()`/`End Sub`。
        /// ★不能包进过程里：`Declare PtrSafe` / `Option Explicit` / `Property` 这些"只能写在
        ///   模块级"的关键字包进 Sub 就变成过程体内，会被 KeywordAllowed 正确滤掉 —— 那样
        ///   测的是另一件事。空过程只是给 CleanupScratch 认的标记。
        /// </summary>
        private static void RebuildScaffoldTopLevel(App app, VbeBackend be, int at, string line)
        {
            app.RunOnUi(() =>
            {
                be.ForceReloadModule();
                int cnt = be.LineCount();
                if (cnt >= at) be.DeleteLines(at, cnt - at + 1);

                // ★第一行留空：下面 `eval` 读的是 `at+1`（与 RebuildScaffold 同一口径，
                //   那套脚手架的 at+1 就是用例行）。用例行必须也落在 at+1，否则断言
                //   读的是别人那一行 —— 第一版没留这个空行，结果 "tr" 敲进了
                //   `Sub VbcEnter()` 里面（日志里写着 `Sub VbcEnter(tr)`）。
                var body = new List<string> { "", line, "Sub VbcEnter()", "End Sub" };
                for (int i = body.Count - 1; i >= 0; i--) be.InsertLines(at, body[i]);

                be.SetCaret(at + 1, line.Length + 1);
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
        /// ★Ctrl 的 down/up 同样是注入的、会被 LLKHF_INJECTED 闸滤掉，但它们照样进系统，
        ///   所以目标键到达时 Down(VK_CONTROL) 为真。目标键经 AcceptInjectedVk 单点放行。
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

        /// <summary>
        /// 本次探针统一走哪条注入路径：默认 Unicode。
        /// ★v181：不再按输入法状态挑路 —— 模式位**读回来不可信**：明明在中文档，
        ///   `IME_CMODE_NATIVE` 却报不置位，照着走真 VK 打进去照样被组字吃掉
        ///   （实测 `For i` 变成 `    For热水若哦…`）。Unicode 注入与输入法无关，
        ///   永远能落字；缺的只有"钩子看到按键"，由 NoteInjectedTyping 补记。
        ///   ★要验真 VK 那条路时置 `VBECOMPLETE_PROBE_VK=1`（前提是自己确认输入法是英文）。
        /// </summary>
        private static bool _imeUnicodePath = !Settings.Flag("PROBE_VK", false);

        /// <summary>
        /// 逐字符"打字"：英数档走真 vk（和人手一致）；中文档走 Unicode 注入
        /// （否则字母被拿去组字，敲出来是 `Fori= 1 To3`，失败原因会伪装成"程序坏了"）。
        /// 返回实际用的方式并写进日志。
        /// ★Unicode 注入没有 VK，钩子按设计看不到注入键 —— 必须靠 NoteInjectedTyping
        ///   补记"刚敲了一个字"，否则 RequireKeyTrigger 判成"不是打字"，候选窗一个都不弹。
        /// </summary>
        private static bool TypeTyping(App app, string s, int delayMs, out string how)
        {
            bool alnum = !_imeUnicodePath;
            how = alnum ? "vk" : "unicode";
            if (alnum) return TypeText(app, s, delayMs);
            foreach (char ch in s)
            {
                Native.SendUnicodeChar(ch);
                // ★Unicode 事件没有 VK，钩子看不到 ⇒ 顺手把"刚敲了一个字"补记进去，
                //   否则 RequireKeyTrigger 判成"不是打字"，候选窗一个都不弹（假红）。
                app.NoteInjectedTyping(ch == ' ' || ch == '\t');
                Thread.Sleep(delayMs);
            }
            return true;
        }

        /// <summary>
        /// 逐字符用例的前置：定下这一轮走哪条注入路径（默认 Unicode），并把输入法现状写进日志。
        /// ★v181：以前切不过去只写一句"⚠️ 切不过去"然后照跑 —— 结果 37 条"该弹"全红，
        ///   看起来像程序坏了，其实一个字都没验到（中文档下注入只能走 Unicode，钩子不理）。
        /// ★只有真要验 VK 路径（PROBE_VK=1）时才去动输入法：改用户的输入法状态是有代价的，
        ///   Unicode 这条路根本不需要它。
        /// </summary>
        private static bool PrepareImeForTyping(Action<string> log, out int mode0, out bool usedShift)
        {
            IntPtr win = Native.FocusedVbeWindow();
            mode0 = Native.ImeConversionMode(win);
            usedShift = false;
            _imeUnicodePath = !Settings.Flag("PROBE_VK", false);
            bool alnum = Native.ImeAlphanumeric(win);
            if (!_imeUnicodePath && !alnum) alnum = Native.ImeForceAlnum(win);
            if (!_imeUnicodePath && !alnum)
            {
                // IMM32 两条通道都没动它：发一个真 Shift（微软拼音 / 搜狗默认拿它切中英）
                Native.SendVk(Native.VK_SHIFT);
                usedShift = true;
                for (int i = 0; i < 10 && !alnum; i++)
                {
                    Thread.Sleep(60);
                    alnum = Native.ImeAlphanumeric(win);
                }
            }
            log("IME 英数模式（仅供参考，不决定注入方式）= " + alnum +
                "（起始 conv=0x" + mode0.ToString("X2") + "，跑完恢复；" +
                "本轮注入走 " + (_imeUnicodePath ? "Unicode + 补记按键" : "真 VK") + "）");
            return alnum;
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
    }
}
