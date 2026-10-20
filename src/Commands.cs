// 五个"改一行"的命令 —— Ctrl+D / Ctrl+K / Ctrl+/ / Ctrl+Shift+Enter / Ctrl+Shift+↑↓。
//
// ★为什么单独一个文件：这些是【会改用户代码】的动作，必须满足两条：
//   1. 同一语义只有一份实现 —— 钩子（App）和真机验证（Program --hktest）走同一批函数，
//      否则"验证通过"证明不了"按键时用的那条路是好的"（v131c 之前就是这样，
//      验证和实际各写一份，结果光标错了两轮都没人发现）。
//   2. 返回值说清楚"到底做成了没有" —— 钩子靠它决定要不要把按键还回去。
//      旧写法是 void：拿不到快照就直接 return，键被吞掉、屏幕上一动不动，
//      用户只能得出"快捷键失效了"这个结论（v131c 用户报的现象）。
//
// 全部只在 UI 线程调用（内部会走 COM）。

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace VbeComplete
{
    internal static class Commands
    {
        private static readonly Regex ReDeclHead =
            new Regex(@"^(?:Dim|Const|Static|Private|Public|Global|ReDim)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReProcHead =
            new Regex(@"^(?:Sub|Function|Property)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>Ctrl+Shift+Enter：在当前行上方插一行，缩进跟着当前行走。</summary>
        public static bool InsertLineAbove(VbeBackend be)
        {
            var snap = be.Snapshot();
            if (snap == null) return false;
            string indent = LeadingWs(snap.Text ?? "");
            if (!be.InsertLines(snap.Line, indent)) return false;
            be.SetCaret(snap.Line, indent.Length + 1);
            return true;
        }

        /// <summary>Ctrl+D：复制当前行到下一行，光标跟到副本上。</summary>
        public static bool DuplicateLine(VbeBackend be)
        {
            var snap = be.Snapshot();
            if (snap == null) return false;
            string line = Rules.Rstrip(snap.Text ?? "");
            if (!be.InsertLines(snap.Line + 1, line)) return false;
            be.SetCaret(snap.Line + 1, CharColOf(snap, line));
            return true;
        }

        /// <summary>
        /// Ctrl+K：删掉当前整行，光标落到【上一行】的行尾。
        /// 删第 1 行 -> 新首行首列；模块只剩一行 -> 清空该行，照样算成功。
        /// </summary>
        public static bool DeleteLine(VbeBackend be)
        {
            var snap = be.Snapshot();
            if (snap == null) return false;
            int count = be.LineCount();
            if (count <= 0) return false;

            if (count == 1) return be.ReplaceLineRaw(1, "");

            if (!be.DeleteLines(snap.Line, 1)) return false;

            if (snap.Line <= 1) be.SetCaret(1, 1);
            else be.CaretToLineEnd(snap.Line - 1);
            return true;
        }

        /// <summary>Ctrl+/：整行注释开关。有选区就逐行处理。</summary>
        public static bool ToggleComment(VbeBackend be)
        {
            var snap = be.Snapshot();
            if (snap == null) return false;

            int from = snap.Line, to = snap.Line;
            int sl, sc, el, ec;
            if (be.TryGetSelection(out sl, out sc, out el, out ec))
            {
                from = Math.Min(sl, el);
                to = Math.Max(sl, el);
                if (ec == 1 && to > from) to--;
            }

            var lines = new List<string>();
            for (int i = from; i <= to; i++)
            {
                string t = be.LineAt(i);
                if (t == null) break;
                lines.Add(Rules.Rstrip(t));
            }
            if (lines.Count == 0) return false;

            bool allCommented = true;
            foreach (var l in lines)
            {
                string tr = l.TrimStart();
                if (tr.Length == 0) continue;
                if (!tr.StartsWith("'", StringComparison.Ordinal)) { allCommented = false; break; }
            }

            bool wrote = false;
            for (int i = 0; i < lines.Count; i++)
            {
                string l = lines[i];
                string nl;
                if (allCommented)
                {
                    int idx = l.IndexOf('\'');
                    nl = idx >= 0 ? l.Remove(idx, 1) : l;
                }
                else
                {
                    int indent = LeadingWs(l).Length;
                    nl = l.Substring(0, indent) + "'" + l.Substring(indent);
                }
                if (be.ReplaceLineRaw(from + i, nl)) wrote = true;
            }

            // ★v131c：`CodeModule.ReplaceLine` 写完会把光标推到【下一行】（真机实测：
            //   在第 10 行 ReplaceLine，GetSelection 变成 (11,1)）。不还原的话，
            //   连按两次 Ctrl+/，第二次读到的已经是下一行 ⇒ 去注释的是下一行，
            //   用户看到的就是"快捷键乱来 / 像失效了"。
            if (wrote)
            {
                string here = be.LineAt(snap.Line) ?? "";
                be.SetCaret(snap.Line, Rules.CharColFromDisp(here, snap.CaretCol, 4));
            }
            return wrote;
        }

        /// <summary>
        /// Ctrl+Shift+↑：落到本过程里【最后一个声明行】的行尾；
        /// Ctrl+Shift+↓：跳到过程收尾的空白行。
        /// </summary>
        public static bool CtrlNavJump(VbeBackend be, int dir)
        {
            var snap = be.Snapshot();
            if (snap == null) return false;
            string text = be.ModuleText();
            if (text == null) return false;

            var lines = Parser.SplitLines(text);
            var procs = Parser.ExtractProcedures(text);
            var proc = Parser.ProcOfLine(procs, snap.Line);
            int start = proc?.StartLine ?? 1;
            int end = proc?.EndLine ?? lines.Length;

            if (dir < 0)
            {
                int target = -1;
                for (int i = snap.Line - 1; i >= start; i--)
                {
                    if (i - 1 < 0 || i - 1 >= lines.Length) break;
                    string code = Parser.CodeOnly(lines[i - 1]).Trim();
                    if (code.Length == 0) continue;
                    if (ReDeclHead.IsMatch(code)) { target = i; break; }
                    if (ReProcHead.IsMatch(code)) break;
                }
                if (target > 0) return be.CaretToLineEnd(target);
                return false;   // 上面没有声明行：把键还回去，别让 Ctrl+Shift+↑ 变成"什么都不发生"
            }

            for (int i = snap.Line + 1; i <= end; i++)
            {
                if (i - 1 < 0 || i - 1 >= lines.Length) break;
                string code = Parser.CodeOnly(lines[i - 1]).Trim();
                if (code.Length == 0) return be.SetCaret(i, 1);
            }
            // 没有空白行：落到过程末尾
            if (end >= 1) return be.CaretToLineEnd(Math.Max(1, end - 1));
            return false;
        }

        /// <summary>把"显示列"的光标位置折成同一行靠左的合理字符列。</summary>
        private static int CharColOf(Snapshot s, string line)
        {
            int n = Math.Min(line.Length, Math.Max(0, s.CaretCol - 1));
            return n + 1;
        }

        internal static string LeadingWs(string s)
        {
            int i = 0;
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t')) i++;
            return s.Substring(0, i);
        }
    }
}
