// 模糊匹配：哪几个候选命中、命中在哪几位、排第几档（只做字符串，不懂 VBA）。
// ★整块连续优先（arr → dataArr），否则逐字符贪心、优先落在词边界（ds → dataSheet）。
// ★贪心必须带反向可行性上界，否则"明明是子序列"也会判成不匹配。
// ★命中位置是逐字符下标集合（UI 拿它逐字符上色），不是区间。

using System;
using System.Collections.Generic;

namespace VbeComplete
{
    /// <summary>一次匹配的结果。Positions 是命中字符的下标（升序）。</summary>
    internal struct FuzzyHit
    {
        public int[] Positions;
        public int Kind;      // 4 完全相等 / 3 前缀 / 2 词首缩略 / 1 连续块 / 0 分散
        public int First;     // 首个命中下标
        public int Span;      // 首尾跨度（-1 表示只有一个字符）
        public int Run;       // 最长连续段长度（判断"是不是真在打这个名字"用）
        /// <summary>
        /// 命中字符里落在词边界（首字母 / 下划线后 / 驼峰处）的**个数**。
        /// ★给排序打分用：VSCode fuzzyScore 的 START_OF_WORD_MATCH，一个 +23。
        ///   与 Kind 不同 —— Kind 是"全部都在词首才算缩略"，这里是逐字符计数，
        ///   所以 `ds → dataSheet` 拿 2，`dl → dataLabel` 也拿 2，可以细比。
        /// </summary>
        public int WordStarts;
        /// <summary>
        /// 命中的字符里，大小写与输入**完全一致**的个数（VSCode UPPER_CASE_MATCH）。
        /// ★打 `MsgB` 时 `MsgBox` 拿满、`msgBox` 拿不满 —— 大小写本身就是用户的意图。
        /// </summary>
        public int CaseHits;
    }

    internal static class Fuzzy
    {
        /// <summary>
        /// 名字里的"词边界"下标集合：首字符 / 下划线之后 / 小写→大写处（驼峰）。
        /// ★必须用原始 name，不能先小写 —— 大小写切换信息一丢就认不出驼峰词首。
        /// ★v190：只留给诊断 / 自检用。匹配热路径**不再建这个集合** —— 每个候选
        ///   一次 HashSet 分配，一个键要几百次，全是 GC 压力（见 IsBoundaryAt）。
        /// </summary>
        public static HashSet<int> WordBoundaries(string name)
        {
            var b = new HashSet<int>();
            if (string.IsNullOrEmpty(name)) return b;
            b.Add(0);
            char prev = name[0];
            for (int i = 1; i < name.Length; i++)
            {
                char ch = name[i];
                if (prev == '_' || (char.IsLower(prev) && char.IsUpper(ch))) b.Add(i);
                prev = ch;
            }
            return b;
        }

        /// <summary>
        /// 第 i 位是不是词边界。等价于 <see cref="WordBoundaries"/> 的 Contains，
        /// 但**不分配**（O(1) 只看前一个字符）。
        /// </summary>
        public static bool IsBoundaryAt(string name, int i)
        {
            if (string.IsNullOrEmpty(name) || i < 0 || i >= name.Length) return false;
            if (i == 0) return true;
            char prev = name[i - 1];
            return prev == '_' || (char.IsLower(prev) && char.IsUpper(name[i]));
        }

        /// <summary>
        /// 小写化，但**保证长度不变**：返回的下标要拿去索引原始名字（高亮 / Kind / Span）。
        /// char.ToLowerInvariant 逐字符转换、一定不变长；写出来是为了让这条约束有个落点。
        /// </summary>
        public static string LowerKeepLen(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var buf = new char[s.Length];
            for (int i = 0; i < s.Length; i++) buf[i] = char.ToLowerInvariant(s[i]);
            return new string(buf);
        }

        /// <summary>
        /// 无法整块连续命中时逐字符挑位置（档位越小越优先，同档取靠前的）：
        /// 0 = 既接上一个命中又在词边界、1 = 词边界（缩略型 ds→dataSheet）、
        /// 2 = 紧接上一个命中、3 = 其它。反向 maxpos 保证"真是子序列就一定凑得出"。
        /// ★v190：直接在**原始 name** 上做大小写无关比较，不再先造一份小写副本、
        ///   也不再建词边界的 HashSet —— 这两样是每个候选一次分配，打字时每秒几百次。
        /// </summary>
        public static int[] GreedyPositions(string name, string query)
        {
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(query)) return null;
            int n = name.Length, k = query.Length;
            if (k > n) return null;

            // 反向可行性上界：maxpos[t] = query[t] 最靠右的合法位置
            var maxpos = new int[k];
            int limit = n - 1;
            for (int t = k - 1; t >= 0; t--)
            {
                if (limit < 0) return null;
                int j = LastIndexOfIgnoreCase(name, query[t], limit);
                if (j < 0) return null;
                maxpos[t] = j;
                limit = j - 1;
            }

            var pos = new int[k];
            int i = 0;
            for (int t = 0; t < k; t++)
            {
                int hi = maxpos[t];
                int bestJ = -1, bestKey = int.MaxValue;
                for (int j = i; j <= hi; j++)
                {
                    if (!SameChar(name[j], query[t])) continue;
                    bool isB = IsBoundaryAt(name, j);
                    int key = (j == i && isB) ? 0 : isB ? 1 : (j == i) ? 2 : 3;
                    if (bestJ < 0 || key < bestKey || (key == bestKey && j < bestJ))
                    {
                        bestKey = key; bestJ = j;
                    }
                }
                if (bestJ < 0) return null;
                pos[t] = bestJ;
                i = bestJ + 1;
            }
            return pos;
        }

        /// <summary>大小写无关的"这两个字符算不算同一个"。中文等无大小写的字符走原值比较。</summary>
        private static bool SameChar(char a, char b)
        {
            if (a == b) return true;
            return char.ToLowerInvariant(a) == char.ToLowerInvariant(b);
        }

        /// <summary>从 <paramref name="limit"/> 往前找 query 字符（大小写无关），找不到返回 -1。</summary>
        private static int LastIndexOfIgnoreCase(string name, char q, int limit)
        {
            for (int j = Math.Min(limit, name.Length - 1); j >= 0; j--)
                if (SameChar(name[j], q)) return j;
            return -1;
        }

        /// <summary>命中下标里最长的那段连续长度。</summary>
        public static int LongestRun(int[] positions)
        {
            if (positions == null || positions.Length == 0) return 0;
            int best = 1, cur = 1;
            for (int i = 1; i < positions.Length; i++)
            {
                cur = positions[i] == positions[i - 1] + 1 ? cur + 1 : 1;
                if (cur > best) best = cur;
            }
            return best;
        }

        /// <summary>
        /// 模糊匹配。query 为空 = 全都要（手动 Ctrl+Space），此时 Kind=0、Positions 为空。
        /// Kind 越大越像用户要的：4 完全相等 > 3 前缀 > 2 词首缩略 > 1 连续块 > 0 分散。
        /// ★v190：整条路径**不再分配小写副本与边界 HashSet**（只留命中下标数组）。
        ///   一个键要跑几百个候选，那两样分配正是打字时 p95 抖到十几毫秒的原因。
        /// </summary>
        public static bool Match(string name, string query, out FuzzyHit hit)
        {
            hit = default(FuzzyHit);
            hit.Span = 0;
            if (string.IsNullOrEmpty(name)) return false;
            if (string.IsNullOrEmpty(query))
            {
                hit.Positions = Array.Empty<int>();
                hit.Kind = 0; hit.First = 0; hit.Span = 0; hit.Run = 0;
                hit.WordStarts = 0; hit.CaseHits = 0;
                return true;
            }
            if (query.Length > name.Length) return false;

            bool contiguous = false;
            int[] positions;
            int at = IndexOfIgnoreCase(name, query, 0);
            if (at >= 0)
            {
                // 整块连续：取"最靠前、且尽量落在词边界"的那一次出现
                int best = -1, bestRank = int.MaxValue;
                int scan = at;
                while (scan >= 0)
                {
                    int rank = scan == 0 ? 0 : (IsBoundaryAt(name, scan) ? 1 : 2);
                    if (rank < bestRank || (rank == bestRank && scan < best))
                    {
                        bestRank = rank; best = scan;
                    }
                    scan = IndexOfIgnoreCase(name, query, scan + 1);
                }
                positions = new int[query.Length];
                for (int i = 0; i < query.Length; i++) positions[i] = best + i;
                contiguous = true;
            }
            else
            {
                positions = GreedyPositions(name, query);
                if (positions == null) return false;
            }

            int kind;
            if (name.Length == query.Length && IndexOfIgnoreCase(name, query, 0) == 0) kind = 4;
            else if (at == 0) kind = 3;
            else
            {
                bool allBound = true;
                foreach (int p in positions) if (!IsBoundaryAt(name, p)) { allBound = false; break; }
                kind = allBound ? 2 : (contiguous ? 1 : 0);
            }

            hit.Positions = positions;
            hit.Kind = kind;
            hit.First = positions[0];
            hit.Span = positions[positions.Length - 1] - positions[0] + 1;
            hit.Run = LongestRun(positions);

            // 打分用的两个计数（逐字符，不做"全有全无"的判定）
            int ws = 0, cs = 0;
            for (int i = 0; i < positions.Length; i++)
            {
                int p = positions[i];
                if (IsBoundaryAt(name, p)) ws++;
                // positions[i] 对应 query 的第 i 个字符 —— 比的是**原始大小写**
                if (name[p] == query[i]) cs++;
            }
            hit.WordStarts = ws;
            hit.CaseHits = cs;
            return true;
        }

        /// <summary>从 from 开始找 query（大小写无关），返回首个命中的下标，找不到 -1。</summary>
        private static int IndexOfIgnoreCase(string name, string query, int from)
        {
            int last = name.Length - query.Length;
            for (int i = Math.Max(0, from); i <= last; i++)
            {
                int j = 0;
                while (j < query.Length && SameChar(name[i + j], query[j])) j++;
                if (j == query.Length) return i;
            }
            return -1;
        }
    }
}
