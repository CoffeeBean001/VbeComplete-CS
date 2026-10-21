// 词库：两个**由用户自己维护**的 SQLite 库。
//
//   keywords.db   —— VBA 语言关键字（Dim / If / Then ...）
//   functions.db  —— VBA 函数名 + 类型名（MsgBox / Split / String ...）
//
// ★为什么是两个库而不是一张表：用户明确要求分开 —— 关键字和函数在他的脑子里
//   就是两类东西，改一个不该动到另一个。分开之后每个库只有 1~4 列，拿任何
//   SQLite 客户端（DB Browser for SQLite 之类）打开就能当表格改。
//
// ★绝不与 Python 版共用文件：那是另一个项目。文件名带 `-cs` 后缀，路径在
//   程序目录下，谁也读不到谁。
//
// ★枚举一律不收。用户原话："vba里的枚举太多了，不要作为提示词。"
//   所以种子里没有任何 vb*/xl*/mso* 常量，只有语言关键字、函数、类型。
//
// ★启动时一次性全量读进内存（几十~几千条，读完就不碰磁盘），
//   之后每一个按键的候选匹配都是纯内存操作 —— 编辑词库需要重启程序才生效，
//   这是刻意的：把"文件被别人改了"这种不确定性挡在运行期之外。
//
// 表结构（首次运行自动创建并灌入默认词条；表非空就不再灌，尊重用户的删除）：
//   keywords (word TEXT PRIMARY KEY, wants_space INTEGER, enabled INTEGER)
//   functions(word TEXT PRIMARY KEY, kind TEXT, wants_parens INTEGER, enabled INTEGER)
//
//   want_space / wants_parens 决定"按 Tab 补完之后跟什么"：
//     * 关键字：绝大多数后面要空格（`Dim `），少数不要（`Me` / `True`）
//     * 函数：默认补一对括号；类型名（String / Long）绝不能补括号
//   enabled = 0 表示"这条我知道，但别提示我" —— 删掉也行，留 0 更好（记得住）。

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace VbeComplete
{
    internal sealed class WordLibrary
    {
        private const string KeywordSchema =
            "CREATE TABLE IF NOT EXISTS keywords (" +
            "word TEXT PRIMARY KEY, " +
            "wants_space INTEGER NOT NULL DEFAULT 1, " +
            "enabled INTEGER NOT NULL DEFAULT 1)";

        private const string FunctionSchema =
            "CREATE TABLE IF NOT EXISTS functions (" +
            "word TEXT PRIMARY KEY, " +
            "kind TEXT NOT NULL DEFAULT 'function', " +
            "wants_parens INTEGER NOT NULL DEFAULT 1, " +
            "enabled INTEGER NOT NULL DEFAULT 1)";

        public static string KeywordPath()
        {
            string p = Settings.Raw("KEYWORD_DB");
            return string.IsNullOrEmpty(p) ? Path.Combine(Settings.Home, "keywords-cs.db") : p;
        }

        public static string FunctionPath()
        {
            string p = Settings.Raw("FUNCTION_DB");
            return string.IsNullOrEmpty(p) ? Path.Combine(Settings.Home, "functions-cs.db") : p;
        }

        public int KeywordCount { get; private set; }
        public int FunctionCount { get; private set; }
        public string LastError { get; private set; } = "";

        /// <summary>
        /// 载入两个库到内存。任何一个打不开都**不致命**：那一个退回内置种子，
        /// 另一个照常工作 —— 补全工具不能因为用户手滑删了个文件就变哑巴。
        /// </summary>
        public void Load()
        {
            // 排错开关：词库被改坏 / 想确认"是不是库的问题"时，一行配置退回种子。
            if (Settings.Flag("NO_WORDDB", false))
            {
                VbaBuiltins.ResetToSeed();
                KeywordCount = VbaBuiltins.KeywordCount;
                FunctionCount = VbaBuiltins.FunctionCount;
                Log.Write("NO_WORDDB=1：跳过两个词库，使用内置种子（关键字 {0} / 函数 {1}）",
                    KeywordCount, FunctionCount);
                return;
            }

            var kw = new List<KeyValuePair<string, bool>>();     // (词, 要不要空格)
            var fn = new List<KeyValuePair<string, string>>();   // (词, 档位)
            var parens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            bool kwOk = LoadKeywords(kw);
            bool fnOk = LoadFunctions(fn, parens);

            if (!kwOk)
            {
                foreach (var w in BuiltinTables.Keywords)
                    kw.Add(new KeyValuePair<string, bool>(w, KeywordSeeds.WantsSpace(w)));
            }
            if (!fnOk)
            {
                // 与 SeedFunctions 同序：类型在前（`String` / `Date` 双身份时按类型算）
                foreach (var w in BuiltinTables.TypeNames)
                    fn.Add(new KeyValuePair<string, string>(w, "type"));
                foreach (var w in BuiltinTables.Functions)
                {
                    fn.Add(new KeyValuePair<string, string>(w, "function"));
                    parens.Add(w);
                }
            }

            KeywordCount = kw.Count;
            FunctionCount = fn.Count;
            VbaBuiltins.Apply(kw, fn, parens);
            WriteReadmeOnce();
            Log.Write("词库已载入：关键字 {0} 条（{1}）、函数/类型 {2} 条（{3}）",
                KeywordCount, kwOk ? KeywordPath() : "内置种子", FunctionCount,
                fnOk ? FunctionPath() : "内置种子");
        }

        private bool LoadKeywords(List<KeyValuePair<string, bool>> outp)
        {
            try
            {
                using (var db = new SqliteDb())
                {
                    if (!db.Open(KeywordPath())) { LastError = db.LastError(); return false; }
                    if (!db.Exec(KeywordSchema)) { LastError = db.LastError(); return false; }
                    if (db.Scalar("SELECT COUNT(*) FROM keywords") == 0 && !SeedKeywords(db))
                    {
                        LastError = "keywords 库首次灌种子失败：" + db.LastError();
                        return false;
                    }

                    foreach (var r in db.Query(
                        "SELECT word, wants_space FROM keywords WHERE enabled<>0 ORDER BY word", 0))
                    {
                        string w = r.Length > 0 ? (r[0] ?? "").Trim() : "";
                        if (w.Length == 0) continue;
                        bool ws = r.Length < 2 || r[1] != "0";
                        outp.Add(new KeyValuePair<string, bool>(w, ws));
                    }
                    // ★v151：判据从"启用行数 > 0"改成"表里到底有没有行"。
                    //   原来的写法把"用户把整张表 enabled 全设成 0"（词库说明里
                    //   明确教的用法：不想删就置 0）判成"库加载失败"，于是退回
                    //   内置种子 —— 用户精心禁用的词全部原地复活，日志还写着
                    //   "内置种子"，看起来像库坏了。
                    //   真失败（库损坏 / 被锁）会让 LastError 有内容，那种情况才
                    //   该回落到种子，保证程序仍然可用。
                    long total = db.Scalar("SELECT COUNT(*) FROM keywords");
                    return total > 0 || string.IsNullOrEmpty(db.LastError());
                }
            }
            catch (Exception e) { LastError = "keywords 库不可用：" + e.Message; return false; }
        }

        private bool LoadFunctions(List<KeyValuePair<string, string>> outp, HashSet<string> parens)
        {
            try
            {
                using (var db = new SqliteDb())
                {
                    if (!db.Open(FunctionPath())) { LastError = db.LastError(); return false; }
                    if (!db.Exec(FunctionSchema)) { LastError = db.LastError(); return false; }
                    if (db.Scalar("SELECT COUNT(*) FROM functions") == 0 && !SeedFunctions(db))
                    {
                        LastError = "functions 库首次灌种子失败：" + db.LastError();
                        return false;
                    }

                    foreach (var r in db.Query(
                        "SELECT word, kind, wants_parens FROM functions WHERE enabled<>0 ORDER BY word", 0))
                    {
                        string w = r.Length > 0 ? (r[0] ?? "").Trim() : "";
                        if (w.Length == 0) continue;
                        string kind = r.Length > 1 ? (r[1] ?? "").Trim().ToLowerInvariant() : "function";
                        if (kind.Length == 0) kind = "function";
                        outp.Add(new KeyValuePair<string, string>(w, kind));
                        if (kind != "type" && (r.Length < 3 || r[2] != "0")) parens.Add(w);
                    }
                    long total = db.Scalar("SELECT COUNT(*) FROM functions");
                    return total > 0 || string.IsNullOrEmpty(db.LastError());
                }
            }
            catch (Exception e) { LastError = "functions 库不可用：" + e.Message; return false; }
        }

        /// <summary>
        /// 首次建库时灌关键字种子。返回 false = 没灌成（调用方回落内置种子）。
        /// <para>★v151：以前 BEGIN / 每一条 INSERT / COMMIT 的返回值全丢掉，日志
        /// 照写"灌入 N 条"。只要有一条写失败（库被别的客户端锁着、磁盘满、
        /// 文件只读），表里就是残缺的 0~N 条，而"表非空就不再灌"这条规矩让
        /// 它**永远**补不回来 —— 词库永久缺字，日志却显示成功。</para>
        /// </summary>
        private static bool SeedKeywords(SqliteDb db)
        {
            if (!db.Exec("BEGIN")) return false;
            foreach (var w in BuiltinTables.Keywords)
                if (!db.Write("INSERT OR IGNORE INTO keywords (word, wants_space, enabled) VALUES (?, ?, 1)",
                        w, KeywordSeeds.WantsSpace(w)))
                {
                    db.Exec("ROLLBACK");
                    Log.Write("keywords 灌种子失败（已回滚）：{0}", db.LastError());
                    return false;
                }
            if (!db.Exec("COMMIT"))
            {
                db.Exec("ROLLBACK");
                Log.Write("keywords 灌种子提交失败：{0}", db.LastError());
                return false;
            }
            Log.Write("keywords 库首次创建：灌入 {0} 条关键字", BuiltinTables.Keywords.Length);
            return true;
        }

        /// <summary>灌函数/类型种子；类型在前（理由见下面注释）。返回 false = 没灌成。</summary>
        private static bool SeedFunctions(SqliteDb db)
        {
            if (!db.Exec("BEGIN")) return false;
            // ★先灌类型、后灌函数，顺序不能反。
            //   `String` / `Date` 在 VBA 里**既是类型又是函数**，两边都想要这一行：
            //   按"类型优先"落库（kind='type', wants_parens=0），于是
            //   `Dim x As String` 后面绝不会被补上一对括号 —— 那正是最容易写错的地方。
            //   反过来的顺序会让它们变成 kind='function'，一按 Tab 就得到 `String()`。
            foreach (var w in BuiltinTables.TypeNames)
                if (!db.Write("INSERT OR IGNORE INTO functions (word, kind, wants_parens, enabled) " +
                              "VALUES (?, 'type', 0, 1)", w))
                {
                    db.Exec("ROLLBACK");
                    Log.Write("functions 灌类型种子失败（已回滚）：{0}", db.LastError());
                    return false;
                }
            foreach (var w in BuiltinTables.Functions)
                if (!db.Write("INSERT OR IGNORE INTO functions (word, kind, wants_parens, enabled) " +
                              "VALUES (?, 'function', 1, 1)", w))
                {
                    db.Exec("ROLLBACK");
                    Log.Write("functions 灌函数种子失败（已回滚）：{0}", db.LastError());
                    return false;
                }
            if (!db.Exec("COMMIT"))
            {
                db.Exec("ROLLBACK");
                Log.Write("functions 灌种子提交失败：{0}", db.LastError());
                return false;
            }
            Log.Write("functions 库首次创建：灌入 {0} 条函数 + {1} 条类型",
                BuiltinTables.Functions.Length, BuiltinTables.TypeNames.Length);
            return true;
        }

        /// <summary>
        /// 第一次建库时在旁边留一份说明（纯文本，用户改库时不用猜列是什么）。
        /// 只在文件不存在时写，绝不覆盖用户自己的笔记。
        /// </summary>
        public static void WriteReadmeOnce()
        {
            try
            {
                string p = Path.Combine(Settings.Home, "词库说明.txt");
                if (File.Exists(p)) return;
                var sb = new StringBuilder();
                sb.AppendLine("VbeComplete（C# 版）的词库说明");
                sb.AppendLine("=================================");
                sb.AppendLine();
                sb.AppendLine("两个库都是标准 SQLite 文件，用任何 SQLite 客户端（DB Browser for SQLite、");
                sb.AppendLine("Navicat、VS Code 的 SQLite 插件都行）直接当表格改。改完重启程序生效。");
                sb.AppendLine();
                sb.AppendLine("keywords-cs.db   →  表 keywords");
                sb.AppendLine("    word         关键字本身，例如 Dim / If / Then");
                sb.AppendLine("    wants_space  补全后要不要跟一个空格：1 要，0 不要（Me / True 这类不要）");
                sb.AppendLine("    enabled      1 提示，0 不提示（不想删又想停用它时用 0）");
                sb.AppendLine();
                sb.AppendLine("functions-cs.db  →  表 functions");
                sb.AppendLine("    word         函数名或类型名，例如 MsgBox / Split / String");
                sb.AppendLine("    kind         function（补一对括号）/ type（类型名，绝不补括号）");
                sb.AppendLine("    wants_parens 1 补括号，0 不补；kind=type 时这一列会被忽略");
                sb.AppendLine("    enabled      1 提示，0 不提示");
                sb.AppendLine();
                sb.AppendLine("注意：");
                sb.AppendLine("  * 枚举常量一律不收（vbRed / xlUp 之类）—— 数量太大，提示出来只会干扰。");
                sb.AppendLine("  * 想加自己的常用函数，往 functions 里插一行即可，程序不区分是不是 VBA 内置的。");
                sb.AppendLine("  * 两个库与 Python 版完全无关，文件名各带 -cs，不要相互拷贝。");
                File.WriteAllText(p, sb.ToString(), new UTF8Encoding(true));
            }
            catch { }
        }
    }

    /// <summary>
    /// "这个关键字补完之后要不要跟空格"的**种子判据**。
    /// 词库建好之后以库里的 wants_space 列为准，这里只在首次灌种子时用一次。
    ///
    /// ★直接沿用 BuiltinTables.KeywordsWantSpace 那份人工校准过的清单
    ///   （`Me` / `Next` / `True` 这些不要空格，`Dim` / `If` / `Set` 这些要）。
    ///   别自己再推一套"看它是不是完整值"的规则 —— 两套判据早晚会打架。
    ///   Python 版和这一版在这张表上是一致的，用户已经按它调过手感。
    /// </summary>
    internal static class KeywordSeeds
    {
        private static readonly HashSet<string> WantSpace = new HashSet<string>(
            BuiltinTables.KeywordsWantSpace, StringComparer.OrdinalIgnoreCase);

        public static bool WantsSpace(string word)
            => !string.IsNullOrEmpty(word) && WantSpace.Contains(word);
    }
}
