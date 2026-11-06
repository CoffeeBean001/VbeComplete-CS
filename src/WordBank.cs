// 词库的读写层（给设置界面用；WordLibrary 只负责"读进来变成内存镜像"）。
// ★WordLibrary.Load 有"表空了就灌种子"的语义，绝不能拿来做增删改 —— 用户删到 0 条会被
//   当成库坏了，一重启种子全回来。这里只做 CRUD，出事都在 LastError 里。
// ★改完必须 Reload()，再由调用方 Engine.InvalidatePool()，新词才立刻出现在候选里。

using System;
using System.Collections.Generic;

namespace VbeComplete
{
    /// <summary>词库里的一行。关键字用 WantsSpace，函数/类型用 Kind + WantsParens。</summary>
    internal sealed class WordEntry
    {
        public string Word = "";
        public bool Enabled = true;
        public bool WantsSpace = true;          // 仅关键字库
        public string Kind = "function";        // 仅函数库：function / type
        public bool WantsParens = true;         // 仅函数库

        public bool IsType => string.Equals(Kind, "type", StringComparison.OrdinalIgnoreCase);

        public WordEntry Clone() => new WordEntry
        {
            Word = Word, Enabled = Enabled, WantsSpace = WantsSpace,
            Kind = Kind, WantsParens = WantsParens,
        };
    }

    internal static class WordBank
    {
        public const string Keywords = "kw";
        public const string Functions = "fn";

        public static string PathOf(string bank)
            => bank == Keywords ? WordLibrary.KeywordPath() : WordLibrary.FunctionPath();

        // ------------------------------------------------------------------
        // 读
        // ------------------------------------------------------------------

        /// <summary>读出整个库（**含** enabled=0 的行 —— 界面上要能把它们重新勾上）。</summary>
        public static List<WordEntry> All(string bank, out string err)
        {
            err = "";
            var list = new List<WordEntry>();
            try
            {
                using (var db = new SqliteDb())
                {
                    string path = PathOf(bank);
                    if (!db.Open(path)) { err = db.LastError(); return list; }
                    string ddl = bank == Keywords ? WordLibrary.KeywordSchema
                                                  : WordLibrary.FunctionSchema;
                    if (!db.Exec(ddl)) { err = db.LastError(); return list; }

                    string sql = bank == Keywords
                        ? "SELECT word, wants_space, enabled FROM keywords ORDER BY word"
                        : "SELECT word, kind, wants_parens, enabled FROM functions ORDER BY word";
                    foreach (var r in db.Query(sql, 0))
                    {
                        var e = new WordEntry { Word = (r.Length > 0 ? r[0] : "").Trim() };
                        if (e.Word.Length == 0) continue;
                        if (bank == Keywords)
                        {
                            e.WantsSpace = r.Length < 2 || r[1] != "0";
                            e.Enabled = r.Length < 3 || r[2] != "0";
                        }
                        else
                        {
                            e.Kind = r.Length > 1 && r[1].Length > 0 ? r[1].ToLowerInvariant() : "function";
                            e.WantsParens = r.Length < 3 || r[2] != "0";
                            e.Enabled = r.Length < 4 || r[3] != "0";
                        }
                        list.Add(e);
                    }
                    if (!string.IsNullOrEmpty(db.LastError())) err = db.LastError();
                }
            }
            catch (Exception e) { err = e.Message; }
            return list;
        }

        /// <summary>库里有几行、其中几行启用着（状态栏用；比 All() 便宜）。</summary>
        public static bool Count(string bank, out int total, out int on, out string err)
        {
            total = on = 0;
            err = "";
            try
            {
                using (var db = new SqliteDb())
                {
                    if (!db.Open(PathOf(bank))) { err = db.LastError(); return false; }
                    string t = bank == Keywords ? "keywords" : "functions";
                    if (!db.Exec(bank == Keywords ? WordLibrary.KeywordSchema
                                                  : WordLibrary.FunctionSchema))
                    { err = db.LastError(); return false; }
                    total = (int)db.Scalar("SELECT COUNT(*) FROM " + t);
                    on = (int)db.Scalar("SELECT COUNT(*) FROM " + t + " WHERE enabled<>0");
                    return true;
                }
            }
            catch (Exception e) { err = e.Message; return false; }
        }

        // ------------------------------------------------------------------
        // 写
        // ------------------------------------------------------------------

        /// <summary>
        /// 插入或覆盖一行。<paramref name="oldWord"/> 非空且与 Word 不同 = 改名
        /// （主键是 word，只能删了再插；同一个事务里做，失败就回滚）。
        /// 返回 true = 写成功。
        /// </summary>
        public static bool Save(string bank, WordEntry e, string oldWord, out string err)
        {
            err = "";
            if (e == null || string.IsNullOrWhiteSpace(e.Word))
            {
                err = "词条是空的";
                return false;
            }
            string word = e.Word.Trim();
            string old = string.IsNullOrWhiteSpace(oldWord) ? null : oldWord.Trim();
            try
            {
                using (var db = new SqliteDb())
                {
                    if (!db.Open(PathOf(bank))) { err = db.LastError(); return false; }
                    if (!db.Exec(bank == Keywords ? WordLibrary.KeywordSchema
                                                  : WordLibrary.FunctionSchema))
                    { err = db.LastError(); return false; }

                    db.Exec("BEGIN");
                    bool ok;
                    if (old != null && !string.Equals(old, word, StringComparison.Ordinal))
                    {
                        ok = db.Write(bank == Keywords
                                ? "DELETE FROM keywords WHERE word = ?"
                                : "DELETE FROM functions WHERE word = ?", old);
                        // 删一行也可能"没这条"（用户手工删过），不算失败
                    }
                    else ok = true;

                    if (ok)
                    {
                        ok = bank == Keywords
                            ? db.Write("INSERT OR REPLACE INTO keywords " +
                                       "(word, wants_space, enabled) VALUES (?, ?, ?)",
                                       word, e.WantsSpace ? 1 : 0, e.Enabled ? 1 : 0)
                            : db.Write("INSERT OR REPLACE INTO functions " +
                                       "(word, kind, wants_parens, enabled) VALUES (?, ?, ?, ?)",
                                       word, e.IsType ? "type" : "function",
                                       e.WantsParens ? 1 : 0, e.Enabled ? 1 : 0);
                    }

                    if (!ok)
                    {
                        err = db.LastError();
                        db.Exec("ROLLBACK");
                        return false;
                    }
                    if (!db.Exec("COMMIT"))
                    {
                        err = db.LastError();
                        db.Exec("ROLLBACK");
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception ex) { err = ex.Message; return false; }
        }

        /// <summary>只改启用状态（列表里勾选 / 取消勾选，一次可能几十行）。</summary>
        public static bool SetEnabled(string bank, IEnumerable<string> words, bool on, out string err)
        {
            err = "";
            try
            {
                using (var db = new SqliteDb())
                {
                    if (!db.Open(PathOf(bank))) { err = db.LastError(); return false; }
                    if (!db.Exec(bank == Keywords ? WordLibrary.KeywordSchema
                                                  : WordLibrary.FunctionSchema))
                    { err = db.LastError(); return false; }
                    string sql = bank == Keywords
                        ? "UPDATE keywords SET enabled = ? WHERE word = ?"
                        : "UPDATE functions SET enabled = ? WHERE word = ?";
                    db.Exec("BEGIN");
                    foreach (var w in words)
                        if (!db.Write(sql, on ? 1 : 0, w))
                        {
                            err = db.LastError();
                            db.Exec("ROLLBACK");
                            return false;
                        }
                    if (!db.Exec("COMMIT")) { err = db.LastError(); db.Exec("ROLLBACK"); return false; }
                    return true;
                }
            }
            catch (Exception e) { err = e.Message; return false; }
        }

        /// <summary>
        /// 删掉若干行。返回**真的**删掉了几条（词不在库里不算错，那一条就是 0）。
        /// ★`Write` 的返回值正是 `sqlite3_changes() > 0`，所以它同时回答了
        ///   "SQL 失败了"和"这条本来就不在"—— 前者会让 err 有内容。
        /// </summary>
        public static int Delete(string bank, IEnumerable<string> words, out string err)
        {
            err = "";
            int n = 0;
            try
            {
                using (var db = new SqliteDb())
                {
                    if (!db.Open(PathOf(bank))) { err = db.LastError(); return 0; }
                    string sql = bank == Keywords
                        ? "DELETE FROM keywords WHERE word = ?"
                        : "DELETE FROM functions WHERE word = ?";
                    db.Exec("BEGIN");
                    foreach (var w in words)
                    {
                        bool one = db.Write(sql, w);
                        if (!one && !string.IsNullOrEmpty(db.LastError()))
                        {
                            err = db.LastError();
                            db.Exec("ROLLBACK");
                            return 0;
                        }
                        if (one) n++;
                    }
                    if (!db.Exec("COMMIT")) { err = db.LastError(); db.Exec("ROLLBACK"); return 0; }
                    return n;
                }
            }
            catch (Exception e) { err = e.Message; return 0; }
        }

        /// <summary>
        /// 把两个库重新灌进内存镜像（VbaBuiltins）。改完词库必须调 —— 否则用户
        /// 加了词、列表里也看得见，候选里却没有，只能靠重启。
        /// </summary>
        public static void Reload() => new WordLibrary().Load();

        /// <summary>
        /// 词条合法性：VBA 标识符（字母/汉字开头，后面字母数字下划线）。
        /// ★不做"是不是 VBA 语法关键字"的校验 —— 用户往库里加自己的常用词是明确支持的，
        ///   这里只挡住"加了也用不上"的东西（带空格、带点、纯数字）。
        /// </summary>
        public static bool IsValidWord(string word, out string why)
        {
            why = "";
            string w = (word ?? "").Trim();
            if (w.Length == 0) { why = "空词条"; return false; }
            if (w.Length > 64) { why = "太长"; return false; }
            foreach (char c in w)
                if (c == ' ' || c == '\t' || c == '.' || c == ',' || c == '(' || c == ')')
                { why = "不能含空格或标点（" + c + "）"; return false; }
            char f = w[0];
            if (!char.IsLetter(f) && f != '_') { why = "首字符必须是字母或下划线"; return false; }
            foreach (char c in w)
                if (!char.IsLetterOrDigit(c) && c != '_') { why = "第 " + (w.IndexOf(c) + 1) + " 个字符非法"; return false; }
            return true;
        }
    }
}
