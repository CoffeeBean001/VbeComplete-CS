// 编辑手感相关的纯判据：块结构识别、缩进推算、自动配对、退格接管。
// ★一行 COM 都不许有：这些函数在键盘钩子线程里跑，碰一次 COM 就可能卡住全系统键盘。
//   只吃"一行文本 + 一个光标列"，因此每条规则都能被自检直接钉住。
// ★与 Python 版刻意不同：缩进单位不按工程反推 TabWidth —— 拟合猜错就把光标落进词中间。

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace VbeComplete
{
    internal static class Editing
    {
        // 可以出现在声明关键字前面的修饰词（`Private Sub Foo()` / `Static Sub`）
        private static readonly string[] ModFirst = { "public", "private", "friend", "static" };

        // 行首是这些词就是"块结构开头"
        private static readonly string[] BlockFirst =
            { "sub", "function", "property", "type", "enum", "for", "do", "while", "with" };

        // 块结构 -> 它欠的收尾。分支（Else/Case）与条件编译（#If）刻意不在表里：
        // 前者没有收尾、后者必须由用户写 `#End If`（我们替它写个裸 `End If` 反而是错的）。
        private static readonly Dictionary<string, string> Closers =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "sub",      "End Sub" },
                { "function", "End Function" },
                { "property", "End Property" },
                { "type",     "End Type" },
                { "enum",     "End Enum" },
                { "if",       "End If" },
                { "with",     "End With" },
                { "select",   "End Select" },
                { "for",      "Next" },
                { "do",       "Loop" },
                { "while",    "Wend" },
            };

        // 收符号表：开符号 -> 收符号。半角与全角（中文标点输入法）都要能配对。
        private static readonly Dictionary<char, char> PairClose =
            new Dictionary<char, char>
            {
                { '(', ')' }, { '"', '"' },
                { '（', '）' }, { '“', '”' },
                // ⚠️ 收符号必须自己占一个键（`')' -> ')'`），就像半角的 `(` 一样：
                //    InsertPair 一进来先查这张表，查不到就直接 return false ——
                //    少了这两行，"敲 `)` 时跨过去"那条路根本走不到。
                { ')', ')' }, { '）', '）' },
            };

        // 退格"一次收一对"认的左符号（左边开、右边闭）
        private static readonly Dictionary<char, char> PairDeleteOpen =
            new Dictionary<char, char>
            {
                { '(', ')' }, { '"', '"' }, { '（', '）' }, { '“', '”' },
            };

        private static readonly Regex ReTailThen =
            new Regex(@"(?:^|\s)then\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReLabel =
            new Regex(@"^\s*(?:[A-Za-z_\u4e00-\u9fff]\w*|\d+)\s*:", RegexOptions.Compiled);

        // ------------------------------------------------------------------
        // 基础
        // ------------------------------------------------------------------

        /// <summary>行首空白（空格 + Tab）。</summary>
        public static string LeadingWs(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            int i = 0;
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t')) i++;
            return s.Substring(0, i);
        }

        /// <summary>行首空白占几列（Tab 按 4 列算）。</summary>
        public static int IndentWidth(string s)
        {
            int n = 0;
            foreach (char c in s ?? "")
            {
                if (c == ' ') n += 1;
                else if (c == '\t') n += 4;
                else break;
            }
            return n;
        }

        /// <summary>去掉字符串与注释之后的"代码形状"部分（行尾空白也去掉）。</summary>
        public static string CodePart(string line)
        {
            if (string.IsNullOrEmpty(line)) return "";
            if (line.Trim().Length == 0) return "";
            return Parser.CodeOnly(line).TrimEnd();
        }

        /// <summary>整行是不是注释（`'...` 或 `Rem ...`）—— 整行没有任何代码。</summary>
        public static bool IsCommentOnly(string line)
        {
            string t = (line ?? "").Trim();
            if (t.Length == 0) return false;
            if (t[0] == '\'') return true;
            string low = t.ToLowerInvariant();
            return low == "rem" || low.StartsWith("rem ") || low.StartsWith("rem\t");
        }

        /// <summary>行标签（`1:` / `Retry:` / `10: x = 1`）—— 惯例顶格，缩进不代表层级。</summary>
        public static bool IsLabelLine(string line)
        {
            string code = CodePart(line);
            if (code.Length == 0) return false;
            return ReLabel.IsMatch(code);
        }

        /// <summary>条件编译指令的第一段（不含 `#`）：`#If VBA7 Then` -> "if"。</summary>
        public static string DirectiveWord(string line)
        {
            string code = CodePart(line);
            if (code.Length == 0) return null;
            string low = Collapse(code).ToLowerInvariant();
            if (!low.StartsWith("#")) return null;
            string rest = low.Substring(1).Trim();
            if (rest.Length == 0) return null;
            return rest.Split(' ')[0];
        }

        private static string Collapse(string s)
        {
            var sb = new StringBuilder(s.Length);
            bool sp = false;
            foreach (char c in s)
            {
                if (c == ' ' || c == '\t')
                {
                    if (!sp) { sb.Append(' '); sp = true; }
                    continue;
                }
                sb.Append(c);
                sp = false;
            }
            return sb.ToString().Trim();
        }

        // ------------------------------------------------------------------
        // 块结构
        // ------------------------------------------------------------------

        /// <summary>
        /// 这一行是哪一种块结构的开头？不是块头返回 null。
        ///   "branch" = Else / ElseIf…Then / Case（缩进一级但**没有收尾**）；
        ///   "directive" = #If / #Else（收尾由用户写 `#End If`）；其余见 Closers。
        /// </summary>
        public static string BlockKind(string line)
        {
            string code = CodePart(line);
            if (code.Length == 0) return null;
            if (code.EndsWith("_", StringComparison.Ordinal)) return null;   // 续行符：语句没写完

            string low = Collapse(code).ToLowerInvariant();
            bool hashed = low.StartsWith("#");
            string[] words = low.Split(' ');
            int i = 0;
            while (i < words.Length && IsModWord(words[i])) i++;
            string first = i < words.Length ? words[i].TrimStart('#') : "";

            if (first == "else" || first == "elseif" || first == "case")
            {
                // `ElseIf x` 没写 Then —— 不算块头
                if (first == "elseif" && !ReTailThen.IsMatch(low)) return null;
                return hashed ? "directive" : "branch";
            }
            if (ReTailThen.IsMatch(low))
            {
                // `If x Then` / `#If VBA7 Then`。单选 If（`If x Then y = 1`）以别的词
                // 结尾，天然被排除。
                return hashed ? "directive" : "if";
            }
            if (first == "select")
            {
                bool ok = words.Length > i + 1 && words[i + 1] == "case";
                if (!ok) return null;
                return hashed ? "directive" : "select";
            }
            foreach (var w in BlockFirst)
                if (w == first) return hashed ? "directive" : first;
            return null;
        }

        private static bool IsModWord(string w)
        {
            foreach (var m in ModFirst) if (m == w) return true;
            return false;
        }

        /// <summary>这一行是不是"块结构的开头"—— 换行后新行该缩进一级。</summary>
        public static bool OpensBlock(string line) => BlockKind(line) != null;

        /// <summary>块头那行该补的收尾；不需要收尾（分支 / 条件编译 / 普通行）返回 null。</summary>
        public static string BlockCloser(string line)
        {
            string kind = BlockKind(line);
            if (kind == null) return null;
            string closer;
            return Closers.TryGetValue(kind, out closer) ? closer : null;
        }

        /// <summary>这一行本身是不是某个块的收尾（`End If` / `Next i` / `Loop`…）。</summary>
        public static bool IsBlockCloserLine(string line)
        {
            foreach (var c in Closers.Values)
                if (IsCloserText(line, c)) return true;
            return false;
        }

        /// <summary>这一行是不是 closer 那一行（`Next  '注释` 也算 Next）。</summary>
        public static bool IsCloserText(string line, string closer)
        {
            if (string.IsNullOrEmpty(closer)) return false;
            string s = CodePart(line).Trim().ToLowerInvariant();
            if (s.Length == 0) return false;
            string want = closer.Trim().ToLowerInvariant();
            return s == want || s.StartsWith(want + " ", StringComparison.Ordinal);
        }

        // ------------------------------------------------------------------
        // 缩进推算
        // ------------------------------------------------------------------

        /// <summary>
        /// 换行后新行该有的行首缩进。
        ///   * 本行是整行注释或空行 -> 往上找第一个"既不是注释也不是空行"的行当参考；
        ///   * 新行 = 参考行的行首空白；参考行是块头再深一级；
        ///   * 连参考行都找不到 -> 用本行自己的空白（= VBE 原生行为，绝不更差）。
        /// </summary>
        public static string NextLineIndent(string curLine, IList<string> above, string unit)
        {
            string reference = null;
            if (!string.IsNullOrEmpty(curLine) && curLine.Trim().Length > 0 && !IsCommentOnly(curLine))
            {
                reference = curLine;
            }
            else if (above != null)
            {
                for (int i = above.Count - 1; i >= 0; i--)
                {
                    string t = above[i];
                    if (!string.IsNullOrEmpty(t) && t.Trim().Length > 0 && !IsCommentOnly(t))
                    {
                        reference = t;
                        break;
                    }
                }
            }
            if (reference == null) return LeadingWs(curLine);
            string baseWs = LeadingWs(reference);
            return OpensBlock(reference) ? baseWs + unit : baseWs;
        }

        /// <summary>
        /// 块【内】分支行该对齐到哪里：返回 (本行应有的行首空白, 新行应有的行首空白)。
        /// 不是分支行、或上方找不到它所属的块头 -> false（调用方维持原行为，绝不猜）。
        /// `Else` / `ElseIf` 与 `If … Then` 齐平；`Case` 比 `Select Case` 深一级。
        /// </summary>
        public static bool BranchAlign(string line, IList<string> above, string unit,
            out string mine, out string next)
        {
            mine = next = null;
            string word = BranchWord(line);
            if (word == null) return false;

            bool hashed = CodePart(line).TrimStart().StartsWith("#", StringComparison.Ordinal);
            string kind = hashed ? "directive" : (word == "case" ? "select" : "if");
            string owner = UnclosedOwnerAbove(above, kind);
            if (owner == null) return false;

            mine = word == "case" ? owner + unit : owner;
            next = mine + unit;
            return true;
        }

        /// <summary>这一行是哪种块内分支：else / elseif / case；不是返回 null。</summary>
        public static string BranchWord(string line)
        {
            string code = CodePart(line);
            if (code.Length == 0) return null;
            if (code.EndsWith("_", StringComparison.Ordinal)) return null;
            string low = Collapse(code).ToLowerInvariant();
            string[] words = low.Split(' ');
            int i = 0;
            while (i < words.Length && IsModWord(words[i])) i++;
            string first = i < words.Length ? words[i].TrimStart('#') : "";
            if (first == "elseif") return ReTailThen.IsMatch(low) ? "elseif" : null;
            if (first == "else" || first == "case") return first;
            return null;
        }

        /// <summary>上方最近一个"还开着"的同类块头的行首空白；没有返回 null。</summary>
        private static string UnclosedOwnerAbove(IList<string> above, string kind)
        {
            string closer = kind == "directive" ? "#End If" : null;
            if (closer == null && !Closers.TryGetValue(kind, out closer)) return null;
            if (above == null) return null;

            int depth = 0;
            for (int i = above.Count - 1; i >= 0; i--)
            {
                string raw = above[i];
                if (string.IsNullOrEmpty(raw) || raw.Trim().Length == 0) continue;

                if (IsCloserText(raw, closer)) { depth++; continue; }
                if (kind == "directive")
                {
                    if (DirectiveWord(raw) != "if") continue;
                }
                else if (!string.Equals(BlockKind(raw), kind, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                if (depth == 0) return LeadingWs(raw);
                depth--;
            }
            return null;
        }

        /// <summary>一级缩进用什么：配置优先；否则跟着工程里的既有 Tab 走；默认 4 空格。</summary>
        public static string IndentUnitFor(IEnumerable<string> lines)
        {
            string raw = Settings.Raw("INDENT_UNIT");
            if (!string.IsNullOrEmpty(raw))
            {
                string low = raw.Trim().ToLowerInvariant();
                if (low == "tab" || low == "\\t" || raw == "\t") return "\t";
                int n;
                if (int.TryParse(low, out n)) return new string(' ', Math.Max(1, Math.Min(16, n)));
            }
            if (lines != null)
            {
                foreach (var t in lines)
                    if (LeadingWs(t).IndexOf('\t') >= 0) return "\t";
            }
            return "    ";
        }

        /// <summary>
        /// 下方到底要不要补 closer（移植 PY v80a/v82/v86 的完整口径）：
        /// 已有本块收尾 → 不补；下方是同类但更浅 → 问 OwnerAboveAt；其余 → 补。
        /// 空行 / 整行注释 / 分支行 / 行标签 / `#` 指令一律跳过（缩进不代表层级）。
        /// ★同级旧块压在下面时也要补 —— 旧口径"同级一律不补"会让新块永远没有 Next。
        /// 不传 kind 时保持旧口径（只有老测试在用）。
        /// </summary>
        public static bool CloserNeeded(IList<string> below, string closer, int baseIndent)
            => CloserNeeded(below, closer, baseIndent, null, null);

        public static bool CloserNeeded(IList<string> below, string closer, int baseIndent,
            string kind, IList<string> above)
        {
            if (string.IsNullOrEmpty(closer) || below == null) return false;

            string first = null;
            foreach (var raw in below)
            {
                if (!string.IsNullOrEmpty(raw) && raw.Trim().Length > 0) { first = raw; break; }
            }

            if (kind == null)
            {
                // ---- 旧口径（v79b）：不看 kind，只看"下方第一个非空行" ----
                foreach (var raw in below)
                {
                    if (string.IsNullOrEmpty(raw) || raw.Trim().Length == 0) continue;
                    if (CodePart(raw).Trim().Length == 0) continue;      // 整行注释

                    int w = IndentWidth(raw);
                    if (w > baseIndent) return false;                    // 块里已经有内容
                    if (IsCloserText(raw, closer)) return false;         // 本块的收尾已在下面
                    if (!IsBlockCloserLine(raw)) return false;           // 同缩进的别的东西 = 块已结束
                }
                return true;
            }

            // ---- 完整口径（v80a）----
            if (OwnCloserBelow(below, closer, kind, baseIndent)) return false;
            if (first != null)
            {
                int w = IndentWidth(first);
                if (IsCloserText(first, closer) && w < baseIndent)
                    return OwnerAboveAt(above, kind, w);
            }
            return true;
        }

        /// <summary>
        /// 下方是不是已经挂着本块自己的收尾（移植 PY `_own_closer_below`）。
        /// 只按"与块头同缩进"的行配平：同类块头 +1、同类收尾 -1，配平到 0 遇到的就是本块的；
        /// 同缩进的别的东西 = 本块早结束，更浅 = 已出这一层，更深的同类另用 deeper 配平。
        /// </summary>
        private static bool OwnCloserBelow(IList<string> below, string closer, string kind, int baseIndent)
        {
            int depth = 0;      // 同缩进的同类块头（本块下面嵌着的同类块）
            int deeper = 0;     // 真正嵌套（缩进更深）的同类块头

            foreach (var raw in below)
            {
                if (string.IsNullOrEmpty(raw) || raw.Trim().Length == 0) continue;
                if (IsCommentOnly(raw) || BranchWord(raw) != null || IndentBlind(raw)) continue;

                int w = IndentWidth(raw);
                if (w < baseIndent) return false;                 // 出了这一层

                if (IsCloserText(raw, closer))
                {
                    if (w == baseIndent)
                    {
                        if (depth == 0) return true;              // 这个收尾就是本块的
                        depth--;
                    }
                    else if (deeper > 0) deeper--;                // 闭合的是更深的那个同类块
                    else return true;                             // 孤零零一个更深的同类收尾：别再补
                    continue;
                }

                if (string.Equals(BlockKind(raw), kind, StringComparison.OrdinalIgnoreCase))
                {
                    if (w == baseIndent) depth++;
                    else deeper++;
                    continue;
                }

                if (w == baseIndent) return false;                // 同缩进的其它东西 -> 本块已结束
            }
            return false;
        }

        /// <summary>
        /// 上方有没有一个【同类、仍未闭合】的块头，缩进正好在 <paramref name="indent"/>？
        /// （移植 Python `_closer_owner_above`）—— 用来判断"下面那个更浅的同类收尾"
        /// 到底是外层块的，还是用户自己写歪的（写歪的再补一条就成了双份收尾）。
        /// </summary>
        private static bool OwnerAboveAt(IList<string> above, string kind, int indent)
        {
            string closer;
            if (!Closers.TryGetValue(kind, out closer)) return false;
            if (above == null) return false;

            int depth = 0;
            for (int i = above.Count - 1; i >= 0; i--)
            {
                string raw = above[i];
                if (string.IsNullOrEmpty(raw) || raw.Trim().Length == 0) continue;

                if (IsCloserText(raw, closer)) { depth++; continue; }
                if (!string.Equals(BlockKind(raw), kind, StringComparison.OrdinalIgnoreCase)) continue;
                if (depth == 0) return IndentWidth(raw) == indent;
                depth--;
            }
            return false;
        }

        /// <summary>这一行的缩进【不代表结构层级】：行标签（`1:`）与 `#` 条件编译指令。</summary>
        private static bool IndentBlind(string line)
            => IsLabelLine(line) || DirectiveWord(line) != null;

        // ------------------------------------------------------------------
        // 钩子线程用的三个"要不要接管"判据
        // ------------------------------------------------------------------

        /// <summary>
        /// 这一行"看起来像块头 / 整行注释"—— 允许只写了半截。
        /// ★只做**宽松**判据，决定"要不要先接管"：钩子线程拿到的是轮询快照，用户敲完最后
        ///   一个字符立刻回车时快照还少一位，严格判据会把回车直接放给 VBE（不补收尾）。
        ///   真判定交给 NewLineBelow 拿真 COM 数据重判，判不过就把回车还回去。
        /// </summary>
        public static bool LooksLikeBlockHeaderPrefix(string line)
        {
            string code = CodePart(line);
            if (code.Length == 0) return IsCommentOnly(line);
            if (code.EndsWith("_", StringComparison.Ordinal)) return false;

            string low = Collapse(code).ToLowerInvariant();
            string[] words = low.Split(' ');
            int i = 0;
            while (i < words.Length && IsModWord(words[i])) i++;
            string first = i < words.Length ? words[i].TrimStart('#') : "";
            switch (first)
            {
                case "if": case "for": case "do": case "while": case "with":
                case "select": case "sub": case "function": case "property":
                case "type": case "enum":
                    return true;
                // ★v184：块内分支行（`Else` / `ElseIf … Then` / `Case …`）也必须算进来。
                //   用户敲完 `Then` 立刻回车，快照往往还少最后一格 ⇒ 严格判据
                //   `EnterIndentWanted` 为 false，只靠这一档"先接管、再拿真数据重判"兜住。
                //   以前这里没有这三个词，于是 `ElseIf` 行整下漏给 VBE 原生：
                //   行不拉齐、新行也不缩进 —— 用户报的"写 ElseIf 再回车不格式化"。
                //   A/B 口子：NO_BRANCH_PREFIX=1 退回 v183 的行为（只为证明这一条判据在起作用）。
                case "else": case "elseif": case "case":
                    return !Settings.Flag("NO_BRANCH_PREFIX", false);
            }
            return false;
        }

        /// <summary>
        /// 这一下回车值不值得由我们接管：只认整行注释与块结构开头，另有三道保守门槛
        /// （只能在行尾、行内不能有 Tab、空行交给 VBE）。
        /// ★caretDisp 是显示列（全角占 2 格、Tab 展开），所以拿"行尾显示列"当靶子比 ——
        ///   全角行上跟字符数比会算错。
        /// </summary>
        public static bool EnterIndentWanted(string line, int caretDisp)
        {
            string t = line ?? "";
            if (t.Trim().Length == 0) return false;
            if (t.IndexOf('\t') >= 0) return false;
            if (caretDisp <= 0) return false;
            if (caretDisp != Rules.DispCol(t, t.Length)) return false;
            return IsCommentOnly(t) || OpensBlock(t);
        }

        /// <summary>
        /// 这一下退格要不要由我们接管 —— 整行只有空白（VBE 回车继承缩进造出来的
        /// `    `）时一次删掉整行。真空行不接管（原生退格结果一样），有内容的行
        /// 更不接管；光标必须停在这行空白的末尾。
        /// </summary>
        public static bool BlankLineDeleteWanted(string line, int caretDisp)
        {
            string t = line ?? "";
            if (t.Length == 0) return false;
            if (t.Trim().Length > 0) return false;
            if (caretDisp <= 0) return false;
            return caretDisp == t.Length + 1 || caretDisp == Rules.DispCol(t, t.Length);
        }

        /// <summary>
        /// 光标夹在一对**空**括号 / 空引号正中间时，给出那一对的下标（0-based）。
        /// caretDisp 先换算成字符列再判 —— 换算只走一条路，两种都试就从精确判据退化成
        /// 并集，会误删。
        /// </summary>
        public static bool PairDeleteSpan(string line, int caretDisp, out int left, out int right)
        {
            left = right = -1;
            string t = line ?? "";
            if (t.Length == 0) return false;
            if (caretDisp <= 0) return false;

            int cc = Rules.CharColFromDisp(t, caretDisp);
            left = cc - 2;
            right = left + 1;
            if (left < 0 || right >= t.Length) { left = right = -1; return false; }

            char close;
            if (PairDeleteOpen.TryGetValue(t[left], out close) && t[right] == close) return true;
            left = right = -1;
            return false;
        }

        // ------------------------------------------------------------------
        // 自动配对
        // ------------------------------------------------------------------

        /// <summary>
        /// 算出自动配对后的结果；null = 一个字符都不用写（右半边已在右边，只把光标跨过去）。
        /// 引号开收同字，左边引号个数为奇数说明这一下是收尾；`(` 右边还有内容时只插左括号，
        /// 除非光标正待在实参/形参表里（那是开新的一层）。
        /// </summary>
        public static string PairInsertion(string line, int col0, char open, char close,
            out int caretOff, out bool paired)
        {
            line = line ?? "";
            if (col0 < 0) col0 = 0;
            if (col0 > line.Length) col0 = line.Length;
            caretOff = col0 + 1;
            paired = false;

            bool isQuote = open == '"' || open == '“';
            if (isQuote)
            {
                if (col0 < line.Length && line[col0] == close) return null;   // 跨过去
                string pre = line.Substring(0, col0);
                bool inside = open == '“'
                    ? CountOf(pre, '“') > CountOf(pre, '”')
                    : CountOf(pre, '"') % 2 == 1;
                if (inside) return line.Substring(0, col0) + close + line.Substring(col0);
                paired = true;
                return line.Substring(0, col0) + open + close + line.Substring(col0);
            }
            if (open == '(' || open == '（')
            {
                if (line.Substring(col0).Trim().Length > 0 && EnclosingCallParen(line, col0) < 0)
                    return line.Substring(0, col0) + open + line.Substring(col0);   // 只插左半边
                paired = true;
                return line.Substring(0, col0) + open + close + line.Substring(col0);
            }
            paired = true;
            return line.Substring(0, col0) + open + close + line.Substring(col0);
        }

        public static char CloseOf(char open)
        {
            char c;
            return PairClose.TryGetValue(open, out c) ? c : '\0';
        }

        private static int CountOf(string s, char c)
        {
            int n = 0;
            foreach (char x in s) if (x == c) n++;
            return n;
        }

        /// <summary>
        /// 光标之前"最内层、未闭合、且左边紧挨标识符"的左括号下标；没有返回 -1。
        /// 这是"括号组是调用/形参表"与"括号组只是运算分组"的分界线：
        ///   `foo(` / `Sub Foo(` -> 是；`x = (a + (` -> 不是（宁可少接管）。
        /// </summary>
        public static int EnclosingCallParen(string line, int col0)
        {
            line = line ?? "";
            int end = Math.Max(0, Math.Min(col0, line.Length));
            if (end <= 0) return -1;

            var stack = new List<int>();
            bool inStr = false;
            for (int i = 0; i < end; i++)
            {
                char ch = line[i];
                if (inStr)
                {
                    if (ch == '"') inStr = false;
                    continue;
                }
                if (ch == '"') { inStr = true; continue; }
                if (ch == '\'') break;                    // 注释开始，后面的不算代码
                if (ch == '(' || ch == '（') stack.Add(i);
                else if (ch == ')' || ch == '）')
                {
                    if (stack.Count > 0) stack.RemoveAt(stack.Count - 1);
                }
            }
            if (stack.Count == 0) return -1;

            int k = stack[stack.Count - 1];
            int j = k - 1;
            while (j >= 0 && (line[j] == ' ' || line[j] == '\t')) j--;
            if (j < 0) return -1;
            char p = line[j];
            if (Rules.IsIdentChar(p) || p == '%' || p == '&' || p == '@' || p == '$' || p == '!')
                return k;
            return -1;
        }

        /// <summary>这一行的代码部分引号是否没闭合（VBE 原生回车会替用户补右引号）。</summary>
        public static bool UnterminatedString(string line)
        {
            string t = line ?? "";
            bool inStr = false;
            int i = 0, n = t.Length;
            while (i < n)
            {
                char c = t[i];
                if (c == '\'' && !inStr) break;
                if (c == '"')
                {
                    if (inStr && i + 1 < n && t[i + 1] == '"') { i += 2; continue; }
                    inStr = !inStr;
                }
                i++;
            }
            return inStr;
        }

        // ------------------------------------------------------------------
        // 其它判据
        // ------------------------------------------------------------------

        /// <summary>这个键按下去"本可能"往代码里写一个字符（数输入法组字用）。</summary>
        // ------------------------------------------------------------------
        // 过程内跳转（Ctrl+↑ / Ctrl+↓）
        // ------------------------------------------------------------------

        /// <summary>
        /// `Ctrl+↑` 认的"声明行"：**只有 `Dim`**。刻意不认 `Const` / `Static` ——
        /// 用户的口径是"最后一个 dim 行"，收宽了就会把光标带到不是 Dim 的那一行。
        /// （Python 版 v110 的 `_RE_V110_DIM` 也是只认 `Dim`。）
        /// </summary>
        private static readonly Regex ReDimDeclHead =
            new Regex(@"^Dim\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// `Ctrl+↑` 的落点：过程（startLine..endLine，1-based 闭区间）里最后一个 `Dim`
        /// 语句的结束行；找不到返回 0（调用方据此把键还回去）。
        /// ★跨续行的 Dim 要整条跨过去，落在 `Dim a, _` 上等于落在半条语句里。
        /// </summary>
        public static int LastDimEndLine(IList<string> lines, int startLine, int endLine)
        {
            if (lines == null || lines.Count == 0) return 0;
            int end = Math.Min(endLine, lines.Count);
            for (int i = end - 1; i > startLine; i--)      // 过程头 / 收尾行本身不扫
            {
                string code = Parser.CodeOnly(lines[i - 1]).Trim();
                if (code.Length == 0) continue;
                if (!ReDimDeclHead.IsMatch(code)) continue;

                int last = i;
                while (last < end && IsContinuation(Parser.CodeOnly(lines[last - 1]).Trim()))
                    last++;
                return last >= end ? 0 : last;            // 续行拖到收尾行：不猜
            }
            return 0;
        }

        /// <summary>
        /// 行尾是续行符 ` _`（下划线本身可以是标识符的一部分，所以要求前面是空白）。
        /// </summary>
        public static bool IsContinuation(string code)
        {
            if (string.IsNullOrEmpty(code) || code.Length < 2) return false;
            if (code[code.Length - 1] != '_') return false;
            char prev = code[code.Length - 2];
            return prev == ' ' || prev == '\t';
        }

        /// <summary>
        /// `Ctrl+↓` 的落点：过程收尾行**上方最后一条非空代码**的行号（1-based）。
        /// 空白行不当落点（落到空行的"行尾"就是行首，对用户毫无意义）；
        /// 过程体为空（收尾行紧挨着过程头）返回 0。
        /// </summary>
        public static int ProcLastCodeLine(IList<string> lines, int startLine, int endLine)
        {
            if (lines == null || lines.Count == 0) return 0;
            int end = Math.Min(endLine, lines.Count);
            int target = end - 1;
            while (target > startLine &&
                   Parser.CodeOnly(lines[target - 1]).Trim().Length == 0)
                target--;
            return target <= startLine ? 0 : target;
        }

        public static bool IsTextKey(int vk)
        {
            if (vk == 0x20) return true;                       // Space
            if (vk >= 0x30 && vk <= 0x39) return true;         // 0-9
            if (vk >= 0x41 && vk <= 0x5A) return true;         // A-Z
            if (vk >= 0x60 && vk <= 0x69) return true;         // 小键盘 0-9
            if (vk >= 0xBA && vk <= 0xC0) return true;         // ;=,-./`
            if (vk >= 0xDB && vk <= 0xDE) return true;         // [\]='
            return false;
        }

        /// <summary>
        /// 注释 / 解除注释之后光标该落在哪个字符列（1-based）。只在"改动发生在光标之前"时
        /// 才调整 —— 否则每注释一次左移一格、解除一次右移一格。
        /// 抽成纯函数是为了能被 --selftest 钉住（真机上几乎走不到这条路）。
        /// </summary>
        public static int AfterCommentCharCol(string oldLine, int oldCharCol, bool allCommented)
        {
            if (oldCharCol < 1) oldCharCol = 1;
            if (!allCommented) return oldCharCol + 1;          // 插在行首，必然在光标之前
            int idx = (oldLine ?? "").IndexOf('\'');
            return idx >= 0 && idx + 1 < oldCharCol ? oldCharCol - 1 : oldCharCol;
        }
    }
}
