// 代码片段库：SQLite 两列表（title 主键 + code），外加标题模糊匹配。
//
// ★为什么不用 Microsoft.Data.Sqlite：那是个 NuGet 包，会把一坨原生 sqlite
//   二进制塞进单文件产物里；而这个程序一直靠"一个 exe 就能跑"。
//   Windows 10 起系统自带 `winsqlite3.dll`（System32），导出的就是标准 SQLite
//   API —— 直接 P/Invoke 它，零依赖、零安装，而且写出来的是**真·SQLite 文件**，
//   与 Python 版的 `snippets.db` 格式完全一致（改天想搬回去也不用转换）。
//
//   ⚠️ 别改成自己造一个 JSON 格式：Python 版那 20KB 的 snippets.db 是用户一条条
//      攒出来的，格式一变就得写迁移，迁移写错就是丢片段。
//
// 表结构与 Python 版逐字一致（snippets.py 顶部 SCHEMA）：
//     CREATE TABLE IF NOT EXISTS snippets (title TEXT PRIMARY KEY, code TEXT NOT NULL)
//   刻意不加 id：标题天然唯一、足以当主键，重命名走 UPDATE ... WHERE title=?。
//
// 本文件不碰 COM、不碰 WinForms —— 全是纯数据操作，能被 --snipetest 单独验证。

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace VbeComplete
{
    /// <summary>一条片段 + 当前搜索词在标题上的命中下标（高亮用）。</summary>
    internal sealed class SnippetHit
    {
        public string Title;
        public string Code;
        public List<int> Hits = new List<int>();
    }

    internal sealed class SnippetStore : IDisposable
    {
        // ---- SQLite 常量 ----
        private const int SQLITE_OK = 0;
        private const int SQLITE_ROW = 100;
        private const int SQLITE_DONE = 101;
        private static readonly IntPtr SQLITE_TRANSIENT = new IntPtr(-1);

        private const string Schema =
            "CREATE TABLE IF NOT EXISTS snippets (title TEXT PRIMARY KEY, code TEXT NOT NULL)";

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl,
            CharSet = CharSet.Unicode)]
        private static extern int sqlite3_open16(string filename, out IntPtr db);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl,
            CharSet = CharSet.Unicode)]
        private static extern int sqlite3_prepare16_v2(IntPtr db, string sql, int nBytes,
            out IntPtr stmt, IntPtr pzTail);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl,
            CharSet = CharSet.Unicode)]
        private static extern int sqlite3_bind_text16(IntPtr stmt, int index, string value,
            int nBytes, IntPtr destructor);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr sqlite3_column_text16(IntPtr stmt, int col);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int sqlite3_column_count(IntPtr stmt);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int sqlite3_step(IntPtr stmt);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int sqlite3_finalize(IntPtr stmt);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int sqlite3_close(IntPtr db);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr sqlite3_errmsg16(IntPtr db);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int sqlite3_changes(IntPtr db);

        private IntPtr _db = IntPtr.Zero;
        private string _path;
        private string _lastError = "";

        public string LastError() => _lastError;
        public string Path_ => _path ?? DbPath();

        /// <summary>
        /// 库文件路径：`SNIPPET_DB`（settings.ini / 环境变量）优先，否则项目文件夹
        /// 下的 `snippets.db` —— 与 Python 版同一套口径，一眼能找到、拷走就跟着走。
        /// </summary>
        public static string DbPath()
        {
            string p = Settings.Raw("SNIPPET_DB");
            if (!string.IsNullOrEmpty(p)) return p;
            return System.IO.Path.Combine(Settings.Home, "snippets.db");
        }

        private bool Open()
        {
            if (_db != IntPtr.Zero) return true;
            _path = DbPath();
            try
            {
                string dir = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(_path));
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

                IntPtr db;
                int rc = sqlite3_open16(_path, out db);
                if (rc != SQLITE_OK || db == IntPtr.Zero)
                {
                    _lastError = "打开片段库失败 rc=" + rc + " 路径=" + _path;
                    Log.Write("{0}", _lastError);
                    if (db != IntPtr.Zero) sqlite3_close(db);
                    return false;
                }
                _db = db;
                if (!Exec(Schema))
                {
                    Log.Write("建表失败：{0}", _lastError);
                }
                return true;
            }
            catch (Exception e)
            {
                // DllNotFound（老系统没有 winsqlite3.dll）也走这里，绝不抛给用户
                _lastError = "片段库不可用：" + e.Message;
                Log.Write("{0}", _lastError);
                return false;
            }
        }

        private string ErrMsg()
        {
            if (_db == IntPtr.Zero) return _lastError;
            try
            {
                IntPtr p = sqlite3_errmsg16(_db);
                return p == IntPtr.Zero ? _lastError : Marshal.PtrToStringUni(p);
            }
            catch { return _lastError; }
        }

        /// <summary>执行一条不带参数、也不读结果的 SQL。</summary>
        private bool Exec(string sql)
        {
            if (_db == IntPtr.Zero) return false;
            IntPtr stmt;
            try
            {
                int rc = sqlite3_prepare16_v2(_db, sql, -1, out stmt, IntPtr.Zero);
                if (rc != SQLITE_OK || stmt == IntPtr.Zero)
                {
                    _lastError = "SQL 准备失败：" + ErrMsg();
                    return false;
                }
                try
                {
                    rc = sqlite3_step(stmt);
                    bool ok = rc == SQLITE_DONE || rc == SQLITE_ROW;
                    if (!ok) _lastError = "SQL 执行失败：" + ErrMsg();
                    return ok;
                }
                finally { sqlite3_finalize(stmt); }
            }
            catch (Exception e)
            {
                _lastError = e.Message;
                return false;
            }
        }

        /// <summary>执行一条带 N 个文本参数的写语句，返回是否真的改到了行。</summary>
        private bool ExecWrite(string sql, params string[] args)
        {
            if (!Open()) return false;
            IntPtr stmt;
            try
            {
                int rc = sqlite3_prepare16_v2(_db, sql, -1, out stmt, IntPtr.Zero);
                if (rc != SQLITE_OK || stmt == IntPtr.Zero)
                {
                    _lastError = "SQL 准备失败：" + ErrMsg();
                    return false;
                }
                try
                {
                    for (int i = 0; i < args.Length; i++)
                    {
                        string v = args[i] ?? "";
                        sqlite3_bind_text16(stmt, i + 1, v, v.Length * 2, SQLITE_TRANSIENT);
                    }
                    rc = sqlite3_step(stmt);
                    if (rc != SQLITE_DONE && rc != SQLITE_ROW)
                    {
                        _lastError = "SQL 执行失败：" + ErrMsg();
                        return false;
                    }
                    return sqlite3_changes(_db) > 0;
                }
                finally { sqlite3_finalize(stmt); }
            }
            catch (Exception e)
            {
                _lastError = e.Message;
                return false;
            }
        }

        private List<string> QueryOneColumn(string sql)
        {
            var list = new List<string>();
            if (!Open()) return list;
            IntPtr stmt;
            try
            {
                int rc = sqlite3_prepare16_v2(_db, sql, -1, out stmt, IntPtr.Zero);
                if (rc != SQLITE_OK || stmt == IntPtr.Zero)
                {
                    _lastError = "SQL 准备失败：" + ErrMsg();
                    return list;
                }
                try
                {
                    while (sqlite3_step(stmt) == SQLITE_ROW)
                    {
                        IntPtr p = sqlite3_column_text16(stmt, 0);
                        list.Add(p == IntPtr.Zero ? "" : (Marshal.PtrToStringUni(p) ?? ""));
                    }
                }
                finally { sqlite3_finalize(stmt); }
            }
            catch (Exception e)
            {
                _lastError = e.Message;
                Log.Write("查片段库失败：{0}", e.Message);
            }
            return list;
        }

        // ------------------------------------------------------------------
        // 读
        // ------------------------------------------------------------------

        /// <summary>全部片段 [(标题, 代码)]，按标题排序（不区分大小写）。</summary>
        public List<KeyValuePair<string, string>> All()
        {
            var rows = new List<KeyValuePair<string, string>>();
            if (!Open()) return rows;
            // ★每次全量读都重新评估：清掉上一次的失败痕迹，本次若真出错会再写上。
            //   不清的话状态栏会一直挂着"几天前"的旧错误。
            _lastError = "";
            IntPtr stmt;
            try
            {
                const string sql = "SELECT title, code FROM snippets ORDER BY title COLLATE NOCASE";
                int rc = sqlite3_prepare16_v2(_db, sql, -1, out stmt, IntPtr.Zero);
                if (rc != SQLITE_OK || stmt == IntPtr.Zero) return rows;
                try
                {
                    while (sqlite3_step(stmt) == SQLITE_ROW)
                    {
                        string t = ColText(stmt, 0);
                        string c = ColText(stmt, 1);
                        rows.Add(new KeyValuePair<string, string>(t, c));
                    }
                }
                finally { sqlite3_finalize(stmt); }
            }
            catch (Exception e)
            {
                _lastError = e.Message;
                Log.Write("读片段库失败：{0}", e.Message);
            }
            return rows;
        }

        private static string ColText(IntPtr stmt, int col)
        {
            IntPtr p = sqlite3_column_text16(stmt, col);
            return p == IntPtr.Zero ? "" : (Marshal.PtrToStringUni(p) ?? "");
        }

        /// <summary>取一条的代码；不存在返回 null。</summary>
        public string Get(string title)
        {
            if (string.IsNullOrEmpty(title)) return null;
            if (!Open()) return null;
            IntPtr stmt;
            try
            {
                const string sql = "SELECT code FROM snippets WHERE title=?";
                int rc = sqlite3_prepare16_v2(_db, sql, -1, out stmt, IntPtr.Zero);
                if (rc != SQLITE_OK || stmt == IntPtr.Zero) return null;
                try
                {
                    sqlite3_bind_text16(stmt, 1, title, title.Length * 2, SQLITE_TRANSIENT);
                    if (sqlite3_step(stmt) == SQLITE_ROW) return ColText(stmt, 0);
                    return null;
                }
                finally { sqlite3_finalize(stmt); }
            }
            catch (Exception e)
            {
                _lastError = e.Message;
                return null;
            }
        }

        public bool Exists(string title) => Get(title) != null;

        public int Count()
        {
            var list = QueryOneColumn("SELECT COUNT(*) FROM snippets");
            int n;
            return list.Count > 0 && int.TryParse(list[0], out n) ? n : 0;
        }

        // ------------------------------------------------------------------
        // 写
        // ------------------------------------------------------------------

        /// <summary>新增。标题已存在时不覆盖（返回 false，由界面提示改走"保存修改"）。</summary>
        public bool Add(string title, string code)
        {
            title = (title ?? "").Trim();
            if (title.Length == 0) return false;
            // INSERT OR IGNORE：标题重复时 sqlite3_changes 为 0 -> 返回 false，
            // 与 Python 版"捕获 IntegrityError 后返回 False"完全同口径。
            return ExecWrite("INSERT OR IGNORE INTO snippets (title, code) VALUES (?, ?)",
                title, code ?? "");
        }

        /// <summary>改代码（标题不变）。返回是否真的改到了行。</summary>
        public bool Update(string title, string code)
        {
            title = (title ?? "").Trim();
            if (title.Length == 0) return false;
            return ExecWrite("UPDATE snippets SET code=? WHERE title=?", code ?? "", title);
        }

        /// <summary>改标题。新标题已被占用时不动任何数据、返回 false。</summary>
        public bool Rename(string oldTitle, string newTitle)
        {
            oldTitle = (oldTitle ?? "").Trim();
            newTitle = (newTitle ?? "").Trim();
            if (oldTitle.Length == 0 || newTitle.Length == 0) return false;
            if (oldTitle == newTitle) return true;
            if (Exists(newTitle)) return false;
            return ExecWrite("UPDATE snippets SET title=? WHERE title=?", newTitle, oldTitle);
        }

        /// <summary>删除一条。返回是否真的删掉了。</summary>
        public bool Delete(string title)
        {
            title = (title ?? "").Trim();
            if (title.Length == 0) return false;
            return ExecWrite("DELETE FROM snippets WHERE title=?", title);
        }

        public void Dispose()
        {
            if (_db != IntPtr.Zero)
            {
                try { sqlite3_close(_db); } catch { }
                _db = IntPtr.Zero;
            }
        }
    }

    // ----------------------------------------------------------------------
    // 标题模糊匹配（纯函数；移植自 Python 版 snippets.py 的 match / search）
    //
    // 三档命中，强度递减：整标题相等 > 前缀 > 任意位置连续子串 > 子序列
    // （`gc` 命中 `GetCode`）。多个词用空格分开，每个词都要命中（AND）。
    //
    // ⚠️ 大小写归一必须保持长度：`"İ".ToLower()` 会变成 2 个字符，一旦拿它当下标
    //    去切原串就整体错位。所以逐字符转小写、长度不变才采用。
    // ----------------------------------------------------------------------
    internal static class SnippetMatch
    {
        private const string BoundaryChars = " _-.[](){}+*/\\,;:=<>\"'";

        private static string LowerKeepLen(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            var sb = new System.Text.StringBuilder(text.Length);
            foreach (char ch in text)
            {
                string low = ch.ToString().ToLowerInvariant();
                sb.Append(low.Length == 1 ? low[0] : ch);
            }
            return sb.ToString();
        }

        private static bool IsBoundary(string title, int idx)
        {
            if (idx <= 0) return true;
            char prev = title[idx - 1], cur = title[idx];
            if (BoundaryChars.IndexOf(prev) >= 0) return true;
            if (char.IsUpper(cur) && char.IsLower(prev)) return true;
            if (char.IsDigit(cur) && !char.IsDigit(prev)) return true;
            return false;
        }

        /// <summary>单个词在标题上的命中情况；没命中返回 null。</summary>
        private static int? TermScore(string termLow, string title, string titleLow,
            List<int> hitsOut)
        {
            int n = termLow.Length;
            if (n <= 0) return null;

            int pos = titleLow.IndexOf(termLow, StringComparison.Ordinal);
            if (pos >= 0)
            {
                for (int i = pos; i < pos + n; i++) hitsOut.Add(i);
                if (titleLow == termLow) return 100 + n * 4;
                if (pos == 0) return 80 + n * 4;
                return 60 + n * 4;
            }

            // 子序列：贪心从左往右扫
            var hits = new List<int>();
            int qi = 0;
            for (int idx = 0; idx < titleLow.Length && qi < n; idx++)
            {
                if (titleLow[idx] != termLow[qi]) continue;
                hits.Add(idx);
                qi++;
            }
            if (qi < n) return null;

            int score = 20 + n * 2;
            foreach (int h in hits) if (IsBoundary(title, h)) score += 3;
            for (int i = 0; i + 1 < hits.Count; i++) if (hits[i + 1] == hits[i] + 1) score += 1;
            hitsOut.AddRange(hits);
            return score;
        }

        /// <summary>query 能否模糊命中 title。命中返回分数，否则 null。</summary>
        public static int? Score(string query, string title, out List<int> hits)
        {
            hits = new List<int>();
            string q = (query ?? "").Trim();
            if (q.Length == 0) return 0;
            title = title ?? "";
            string titleLow = LowerKeepLen(title);
            int total = 0;
            var all = new List<int>();
            foreach (var term in q.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var got = new List<int>();
                int? s = TermScore(LowerKeepLen(term), title, titleLow, got);
                if (s == null) { hits.Clear(); return null; }
                total += s.Value;
                all.AddRange(got);
            }
            all.Sort();
            var uniq = new List<int>();
            foreach (int h in all) if (uniq.Count == 0 || uniq[uniq.Count - 1] != h) uniq.Add(h);
            hits = uniq;
            return total;
        }

        /// <summary>
        /// 按标题模糊过滤。空 query 原样返回全部（保持库里的顺序）；
        /// 否则按"分数高 -> 标题短 -> 字母序"排。
        /// </summary>
        public static List<SnippetHit> Filter(List<KeyValuePair<string, string>> rows, string query)
        {
            var output = new List<SnippetHit>();
            var src = rows ?? new List<KeyValuePair<string, string>>();
            bool empty = string.IsNullOrWhiteSpace(query);

            if (empty)
            {
                foreach (var r in src)
                    output.Add(new SnippetHit { Title = r.Key, Code = r.Value });
                return output;
            }

            var scored = new List<Tuple<int, int, string, SnippetHit>>();
            foreach (var r in src)
            {
                List<int> hits;
                int? s = Score(query, r.Key, out hits);
                if (s == null) continue;
                scored.Add(Tuple.Create(s.Value, (r.Key ?? "").Length,
                    LowerKeepLen(r.Key),
                    new SnippetHit { Title = r.Key, Code = r.Value, Hits = hits }));
            }
            scored.Sort((a, b) =>
            {
                int c = b.Item1.CompareTo(a.Item1); if (c != 0) return c;
                c = a.Item2.CompareTo(b.Item2); if (c != 0) return c;
                return string.CompareOrdinal(a.Item3, b.Item3);
            });
            foreach (var t in scored) output.Add(t.Item4);
            return output;
        }
    }
}
