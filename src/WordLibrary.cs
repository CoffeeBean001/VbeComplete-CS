// 词库：两个由用户自己维护的 SQLite 库（keywords-cs.db / functions-cs.db，绝不与 PY 版共用）。
// 表结构：keywords(word, wants_space, enabled) / functions(word, kind, wants_parens, enabled)。
//   wants_space 决定关键字补完跟空格（`Dim `），wants_parens 决定函数补括号 —— 类型名
//   （String / Long）绝不补；enabled = 0 = "知道但别提示我"。
// ★启动时一次性全量读进内存，之后每个按键都是纯内存匹配；表非空就不再灌种子（尊重用户的删除）。

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace VbeComplete
{
    internal sealed class WordLibrary
    {
        internal const string KeywordSchema =
            "CREATE TABLE IF NOT EXISTS keywords (" +
            "word TEXT PRIMARY KEY, " +
            "wants_space INTEGER NOT NULL DEFAULT 1, " +
            "enabled INTEGER NOT NULL DEFAULT 1)";

        internal const string FunctionSchema =
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
                    // ★判据是"表里到底有没有行"，不是"启用行数 > 0"：用户把整张表
                    //   enabled 全设成 0 是词库说明里教的用法，按"启用数"判会当成库加载
                    //   失败而灌回种子 —— 精心禁用的词全部复活。真失败看 LastError 有没有内容。
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
        /// ★BEGIN / 每条 INSERT / COMMIT 的返回值必须检查：只要有一条写失败，表里就是残缺的
        ///   0~N 条，而"表非空就不再灌"会让它**永远**补不回来（词库永久缺字，日志却写成功）。
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
                sb.AppendLine("**推荐：不用碰这两个文件**。托盘右键 →「添加关键字 / 函数…」（或「设置…」→ 词库），");
                sb.AppendLine("在窗口里就能增删改，支持一行一个批量粘贴，改完立刻生效、不用重启。");
                sb.AppendLine();
                sb.AppendLine("两个库都是标准 SQLite 文件，想用 DB Browser for SQLite / Navicat / VS Code 的");
                sb.AppendLine("SQLite 插件当表格改也可以。★外部改的只在启动时读一次，改完要重启程序；");
                sb.AppendLine("只有程序里那套界面改完是即时的。");
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
    /// "这个关键字补完之后要不要跟空格"（只在首次灌种子时用，之后以库里的 wants_space 为准）。
    /// ★直接沿用 BuiltinTables 那份人工校准过的清单，别再推一套"看它是不是完整值"的规则
    ///   —— 两套判据早晚打架。
    /// </summary>
    internal static class KeywordSeeds
    {
        private static readonly HashSet<string> WantSpace = new HashSet<string>(
            BuiltinTables.KeywordsWantSpace, StringComparer.OrdinalIgnoreCase);

        public static bool WantsSpace(string word)
            => !string.IsNullOrEmpty(word) && WantSpace.Contains(word);
    }
}
