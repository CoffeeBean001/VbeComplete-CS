// 编辑手感相关的【纯判据】：块结构识别、缩进推算、自动配对、退格接管。
//
// ★为什么单独一个文件、而且一行 COM 都不许有：
//   这些东西要在【键盘钩子线程】里被调用（决定"这一下按键吃不吃"）。钩子线程
//   没有 COM 单元，碰一次 COM 就可能让全系统的键盘卡住（Python 版 v56 的教训）。
//   所以这里全部只吃"一行文本 + 一个光标列"，能被自检直接钉住。
//
// ★与 Python 版（vbe_bridge.py）的分工一一对应：
//   _code_part / _is_comment_only / _block_kind / opens_block / block_closer
//   / _indent_width / _leading_ws / next_line_indent / branch_align
//   / enter_indent_wanted / blank_line_delete_wanted / pair_delete_span
//   / _pair_insertion / _unterminated_string / _indent_unit_for
//
// ★两个"故意不一样"的地方（都是为了少改坏用户代码，见各函数注释）：
//   1. `CloserNeeded` 用的是【保守口径】：只有光标下方再没有"本块之外的东西"
//      才补收尾。Python 版那套按缩进配平的判据更聪明，但分支更多、更容易在
//      边缘情况上多补一条 `End If`（双份收尾 = 直接编译不过）。宁可少补一条
//      （用户自己敲一下就好），也绝不凭空变出一条。
//   2. 缩进单位不做"按工程反推 tabw/wide2"的拟合：Tab 一律按 4 列算。列语义
//      的拟合一旦猜错，光标就会落进词中间（Python 版第 19 条铁律）。

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

        /// <summary>
        /// 块名 -> 收尾文本。格式化（Indent.cs）要反查"这一行是哪个块的收尾"，
        /// 所以必须能读；但外面不许改（拿着去 Add 一个就会影响回车补收尾）。
        /// </summary>
        public static IReadOnlyDictionary<string, string> CloserTable => Closers;

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
        /// 这一行是哪一种块结构的开头？返回规范化类型名；不是块头返回 null。
        ///   * "branch"    —— Else / ElseIf…Then / Case：要缩进一级，但**没有收尾**；
        ///   * "directive" —— #If / #Else：要缩进一级，收尾必须由用户写 `#End If`；
        ///   * 其余 —— Closers 里有对应收尾的那些。
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
        /// 下方到底要不要补 closer（保守口径，见文件头说明）：
        /// 只有"下方没有任何本块之外的东西"才补 —— 也就是说：
        ///   * 已经挂着一个本块的收尾（缩进 ≥ 块头）-> 不补（否则双份收尾，编译不过）；
        ///   * 下面还有更深的东西 -> 块里已经有内容了，不补（补了就插进用户代码中间）；
        ///   * 下面有同缩进的、又不是任何块收尾的东西 -> 本块早结束了，不补。
        /// 空行与整行注释一律跳过（它们不影响结构）。
        /// </summary>
        public static bool CloserNeeded(IList<string> below, string closer, int baseIndent)
        {
            if (string.IsNullOrEmpty(closer) || below == null) return false;
            foreach (var raw in below)
            {
                if (string.IsNullOrEmpty(raw) || raw.Trim().Length == 0) continue;
                if (CodePart(raw).Trim().Length == 0) continue;          // 整行注释

                int w = IndentWidth(raw);
                if (w > baseIndent) return false;                        // 块里已经有内容
                if (IsCloserText(raw, closer)) return false;             // 本块的收尾已在下面
                if (!IsBlockCloserLine(raw)) return false;               // 同缩进的别的东西 = 块已结束
            }
            return true;
        }

        // ------------------------------------------------------------------
        // 钩子线程用的三个"要不要接管"判据
        // ------------------------------------------------------------------

        /// <summary>
        /// 这一下回车值不值得由我们接管。只有两种情形：整行注释、块结构开头；
        /// 另有三道保守门槛（拿不准就不抢回车）：只能在行尾、行内不能有 Tab、
        /// 空行交给 VBE。
        /// ★caretDisp 是 VBE 报的【显示列】（全角占 2 格、Tab 展开）—— 全角行上
        ///   直接拿它跟字符数比会算错，所以这里改成拿"行尾显示列"当靶子、要求完全
        ///   相等。行内含全角时"恰好等于行尾显示列"的位置只有一个，不会误判。
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
        /// 判据只有一条：光标左边那个字符与右边那个字符是不是一对配对符。
        /// caretDisp 是 VBE 的显示列，这里先换算成字符列再判（换算只走一条路，
        /// 绝不"两种都试"—— 两种都试就从精确判据退化成并集，会误删）。
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
        /// 算出自动配对后的结果。返回 null 表示"一个字符都不用写"（右半边已经
        /// 在光标右边 -> 只把光标跨过去）。
        ///   * 引号特殊（开收同字）：先看左边引号的奇偶，奇数说明这一下是收尾；
        ///   * `(`：右边还有内容就只插左括号（现代编辑器惯例），除非光标正待在
        ///     实参表 / 形参表里（那是"开一层新的"，右边那个 `)` 属于外层）。
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
    }
}
