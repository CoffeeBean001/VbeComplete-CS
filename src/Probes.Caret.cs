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
    /// <summary>真机光标探针：补全后光标落点。</summary>
    internal static partial class Program
    {
        /// <summary>
        /// 光标落点验证（真机）：模拟"打到一半 + Tab 补全"，读回 VBE 真实光标列，跟按真行
        /// 算出来的列对数，并把光标左右两截打出来。
        /// ★动的是模块末尾的临时行，finally 里删干净并逐字节比对原文。
        /// </summary>
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
                    int wordEndDisp = Rules.DispCol(inserted, at0 + c.Item3.Length);
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
                        ? Rules.DispCol(actual, idx + 1)          // `)` 之前
                        : Rules.DispCol(actual, at0 + c.Item4.Length);

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
    }
}
