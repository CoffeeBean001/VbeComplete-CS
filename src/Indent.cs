// 一键格式化缩进（Ctrl+Shift+L）的【纯逻辑】—— 跟 Editing.cs 一样，一行 COM 都没有。
//
// ★为什么单独一个文件、而且要"只算不写"：
//   1. 这块最怕错。真机上一旦把用户的代码缩歪，用户只会得出"这工具乱改我代码"；
//      所以算法做成纯函数，自检可以直接把每一行该有的行首空白钉住。
//   2. 写回由 Commands.FormatIndent 负责（那边才知道光标在哪儿、怎么还原）。
//
// ★算法只有一条规则：维护一个"块栈"，一行一行走。
//   * 遇到块头（Sub/If/For/Do/While/With/Select Case/Type/Enum/#If）-> 本行按当前深度，
//     然后把块压栈；
//   * 遇到收尾（End If/Next/Loop/Wend/End With/End Select/End Sub …/#End If）
//     -> 找到栈里那个块，本行**与块头齐平**，然后把它的内层一起弹掉；
//   * 遇到分支（Else/ElseIf/Case/#Else）-> 本行与"它所属的块头"齐平，正文仍然深一级：
//       - Else/ElseIf 与 If 齐平；
//       - Case 比 Select Case 深一级（VBA 的通行写法，也是 VBE 自己的自动缩进口径）；
//   * 续行（行尾 ` _`）比语句首行深一级，且**不动块栈**（语句还没写完）；
//   * 行标签顶格；空行 / 整行注释**一个字符都不动**（它们不改变结构，也不该被"整理"）。
//
// ★刻意不做的事（都是为了"少改坏"）：
//   * 不删行尾空格、不动行内空格、不改变大小写、不重排括号 —— 只改**行首空白**。
//     于是撤销起来是"缩进变了"，而不是"代码被重写了一遍"。
//   * 拿不准的结构（块的收尾丢了、`End If` 找不到 If）一律**不弹栈**：
//     宁可某一段缩进保持现状，也绝不因为一次猜错把后面整篇带偏。

using System;
using System.Collections.Generic;
using System.Text;

namespace VbeComplete
{
    internal static class Indent
    {
        /// <summary>格式化结果。Ws 与输入等长，null 表示"这一行不用改"。</summary>
        public sealed class Result
        {
            public string[] Ws;
            public int Changed;
            /// <summary>扫到最后还剩几层块（≠0 说明源码本身没收干净，只记日志不拦）。</summary>
            public int FinalDepth;
            internal IList<string> Lines;      // 只想据（Want 要拿原行文本比对）
        }

        /// <summary>
        /// 算出 [from, to] 各行应有的行首空白（1-based，含两端）。**深度永远从第 1 行扫起**，
        /// 所以"只格式化选中的几行"也能得到正确的层级。
        /// </summary>
        public static Result Compute(IList<string> lines, string unit, int from, int to)
        {
            var res = new Result
            {
                Ws = new string[lines == null ? 0 : lines.Count],
                Lines = lines,
            };
            if (lines == null || lines.Count == 0) return res;
            if (string.IsNullOrEmpty(unit)) unit = "    ";
            if (from < 1) from = 1;
            if (to > lines.Count) to = lines.Count;

            var stack = new List<string>();
            // ★续行必须**先攒成一条完整语句再判层级**：`If a And _` / `b Then` 这种
            //   多行块头，只看第一行认不出是 If（行尾是 `_`），层级就会整段错开。
            var pending = new StringBuilder();
            var contIdx = new List<int>();

            for (int i = 0; i < lines.Count; i++)
            {
                string code = Editing.CodePart(lines[i] ?? "");
                if (code.Length == 0) continue;          // 空行 / 整行注释：原样，不打断块栈

                if (pending.Length > 0)
                {
                    pending.Append(' ').Append(code);
                    contIdx.Add(i);
                    if (IsContinuation(code)) continue;
                    Flush(res, pending.ToString(), contIdx, stack, from, to, unit);
                    pending.Length = 0;
                    contIdx.Clear();
                    continue;
                }

                if (IsContinuation(code))
                {
                    pending.Append(code);
                    contIdx.Add(i);
                    continue;
                }

                Want(res, i, Step(code, stack), from, to, unit);
            }

            // 文件末尾还挂着半条续行语句：按"普通单行"收尾，绝不吞行
            if (pending.Length > 0) Flush(res, pending.ToString(), contIdx, stack, from, to, unit);

            res.FinalDepth = stack.Count;
            return res;
        }

        /// <summary>一条续行语句收尾：整条语句算一次层级，续行各再深一级。</summary>
        private static void Flush(Result res, string joined, List<int> idx, List<string> stack,
            int from, int to, string unit)
        {
            if (idx.Count == 0) return;
            int level = Step(joined, stack);
            Want(res, idx[0], level, from, to, unit);
            for (int k = 1; k < idx.Count; k++) Want(res, idx[k], level + 1, from, to, unit);
        }

        /// <summary>记下第 i 行（0-based）该有的行首空白；不在范围内 / 已经对了就不记。</summary>
        private static void Want(Result res, int i, int level, int from, int to, string unit)
        {
            if (level < 0) level = 0;
            if (i + 1 < from || i + 1 > to) return;
            string raw = res.Lines[i] ?? "";
            string want = Repeat(unit, level);
            if (string.Equals(Editing.LeadingWs(raw), want, StringComparison.Ordinal)) return;
            res.Ws[i] = want;
            res.Changed++;
        }

        /// <summary>走一行代码（已去掉注释），返回它该有的层级，并顺带维护块栈。</summary>
        private static int Step(string code, List<string> stack)
        {
            // ---- 条件编译：`#If` 也算一层，收尾是用户写的 `#End If` ----
            string dword = code.TrimStart().StartsWith("#", StringComparison.Ordinal)
                ? Editing.DirectiveWord(code) : null;
            if (dword != null)
            {
                if (dword == "if")
                {
                    int lv = stack.Count;
                    stack.Add("directive");
                    return lv;
                }
                if (dword == "else" || dword == "elseif")
                {
                    int idx = LastIndexOf(stack, "directive");
                    return idx >= 0 ? idx : stack.Count;   // 不弹栈：正文仍深一级
                }
                if (dword == "end")
                {
                    int idx = LastIndexOf(stack, "directive");
                    if (idx < 0) return Math.Max(0, stack.Count - 1);
                    PopTo(stack, idx);
                    return idx;
                }
                return stack.Count;                        // #Const 之类
            }

            // ---- 行标签（`1:` / `Retry:`）惯例顶格，且不代表层级 ----
            if (Editing.IsLabelLine(code)) return 0;

            // ---- 收尾 ----
            string closer = CloserKindOf(code);
            if (closer != null)
            {
                int idx = LastIndexOf(stack, closer);
                if (idx < 0) return Math.Max(0, stack.Count - 1);   // 找不到块头：不弹栈
                PopTo(stack, idx);                                  // 连内层（如 Case）一起弹
                return idx;
            }

            // ---- 分支 ----
            string branch = Editing.BranchWord(code);
            if (branch == "case")
            {
                int idx = LastIndexOf(stack, "select");
                if (idx < 0) return Math.Max(0, stack.Count - 1);
                PopTo(stack, idx + 1);          // 上一个 Case 的块体到此结束，Select 留着
                stack.Add("case");
                return idx + 1;                 // Case 比 Select Case 深一级
            }
            if (branch == "else" || branch == "elseif")
            {
                int idx = LastIndexOf(stack, "if");
                return idx >= 0 ? idx : Math.Max(0, stack.Count - 1);
            }

            // ---- 普通行 / 块头 ----
            int level = stack.Count;
            string kind = Editing.BlockKind(code);
            if (kind != null && kind != "branch" && kind != "directive") stack.Add(kind);
            return level;
        }

        /// <summary>这一行是哪个块的收尾（返回块名）；不是收尾返回 null。</summary>
        private static string CloserKindOf(string code)
        {
            foreach (var kv in Editing.CloserTable)
                if (Editing.IsCloserText(code, kv.Value)) return kv.Key;
            return null;
        }

        /// <summary>行尾是续行符 ` _`（下划线本身可以是标识符的一部分，所以要求前面是空白）。</summary>
        public static bool IsContinuation(string code)
        {
            if (string.IsNullOrEmpty(code) || code.Length < 2) return false;
            if (code[code.Length - 1] != '_') return false;
            char prev = code[code.Length - 2];
            return prev == ' ' || prev == '\t';
        }

        private static int LastIndexOf(List<string> stack, string kind)
        {
            for (int i = stack.Count - 1; i >= 0; i--)
                if (string.Equals(stack[i], kind, StringComparison.OrdinalIgnoreCase)) return i;
            return -1;
        }

        /// <summary>只保留栈底 keep 层。</summary>
        private static void PopTo(List<string> stack, int keep)
        {
            if (keep < 0) keep = 0;
            if (keep < stack.Count) stack.RemoveRange(keep, stack.Count - keep);
        }

        private static string Repeat(string unit, int n)
        {
            if (n <= 0) return "";
            if (unit == "\t") return new string('\t', n);
            var sb = new StringBuilder(unit.Length * n);
            for (int i = 0; i < n; i++) sb.Append(unit);
            return sb.ToString();
        }
    }
}
