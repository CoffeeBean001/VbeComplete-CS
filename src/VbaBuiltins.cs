// 内建名字：语言关键字 / 内建函数 / 内建类型。
//
// ★v133 起，这些名字的**事实来源是两个 SQLite 词库**（keywords-cs.db /
//   functions-cs.db，见 WordLibrary.cs）。本类只是它们的内存镜像 ——
//   启动时 Load 一次，之后所有匹配都是纯内存查表。
//
// ★为什么保留一个"内置种子"：词库文件被删掉 / 打不开 / 用户改用便携路径失败时，
//   补全不能因此变成哑巴。种子表 BuiltinTables 只在这时兜底，正常运行永远
//   走词库。用户在库里加的自己的函数，种子表里没有 —— 那是预期行为。
//
// ★一切枚举常量都不收（`vbRed` / `xlUp` / `msoAlignLeft` 这类）。
//   用户原话："vba里的枚举太多了，不要作为提示词。"
//
// ★`VBA_KEYWORDS` 是全局总开关（默认开）。关掉它 = 关键字一条都不提示，
//   函数与类型照旧 —— 有些人的代码风格里关键字靠肌肉记忆，弹出来反而碍事。

using System;
using System.Collections.Generic;

namespace VbeComplete
{
    internal static class VbaBuiltins
    {
        public static readonly bool EnableKeywords = Settings.Flag("VBA_KEYWORDS", true);

        private static readonly HashSet<string> _kw = Mk();
        private static readonly HashSet<string> _kwSpace = Mk();
        private static readonly HashSet<string> _funcs = Mk();
        private static readonly HashSet<string> _types = Mk();
        private static readonly HashSet<string> _parens = Mk();

        private static HashSet<string> Mk() => new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 首次触碰本类时灌入内置种子。★这一步不能省：Parser / Scope / Rules 在
        /// 自检与单元场景下会直接问"这是不是关键字"，那时词库根本没打开过。
        /// </summary>
        static VbaBuiltins()
        {
            ApplySeed();
        }

        private static void ApplySeed()
        {
            _kw.Clear(); _kwSpace.Clear(); _funcs.Clear(); _types.Clear(); _parens.Clear();
            foreach (var n in BuiltinTables.Keywords)
            {
                _kw.Add(n);
                if (KeywordSeeds.WantsSpace(n)) _kwSpace.Add(n);
            }
            foreach (var n in BuiltinTables.Functions)
            {
                _funcs.Add(n);
                _parens.Add(n);
            }
            foreach (var n in BuiltinTables.TypeNames) _types.Add(n);
        }

        /// <summary>
        /// 用词库内容替换内存镜像。由 <see cref="WordLibrary.Load"/> 在启动时调一次。
        /// </summary>
        public static void Apply(IEnumerable<KeyValuePair<string, bool>> keywords,
            IEnumerable<KeyValuePair<string, string>> functions, HashSet<string> wantParens)
        {
            _kw.Clear(); _kwSpace.Clear(); _funcs.Clear(); _types.Clear(); _parens.Clear();

            if (keywords != null)
                foreach (var kv in keywords)
                {
                    if (string.IsNullOrEmpty(kv.Key)) continue;
                    _kw.Add(kv.Key);
                    if (kv.Value) _kwSpace.Add(kv.Key);
                }

            if (functions != null)
                foreach (var kv in functions)
                {
                    string w = kv.Key;
                    if (string.IsNullOrEmpty(w)) continue;
                    if (string.Equals(kv.Value, "type", StringComparison.OrdinalIgnoreCase)) _types.Add(w);
                    else _funcs.Add(w);
                }

            if (wantParens != null) foreach (var w in wantParens) if (!string.IsNullOrEmpty(w)) _parens.Add(w);
        }

        public static bool IsBuiltinFunction(string name) =>
            !string.IsNullOrEmpty(name) && _funcs.Contains(name);

        public static bool IsBuiltinType(string name) =>
            !string.IsNullOrEmpty(name) && _types.Contains(name);

        public static bool IsKeyword(string name) =>
            !string.IsNullOrEmpty(name) && _kw.Contains(name);

        /// <summary>这个词算不算"内建词汇"（排序时用来把用户自己的名字排前面）。</summary>
        public static bool IsBuiltin(string name) =>
            IsKeyword(name) || IsBuiltinFunction(name) || IsBuiltinType(name);

        /// <summary>补全后要不要跟一个空格（以词库的 wants_space 为准）。</summary>
        public static bool KeywordWantsSpace(string name) =>
            !string.IsNullOrEmpty(name) && _kwSpace.Contains(name);

        /// <summary>
        /// 是不是"该补一对括号"的可调用名字。
        /// 两道判据问的是不同问题，不能合成一条：
        ///   ① 词库 wants_parens —— 挡住类型名与不可调用语句；
        ///   ② 类型排除 —— 挡住 `String` / `Date` 这类"既是类型又是函数"的双身份名字
        ///      （在 `Dim x As String` 的位置补括号就是语法错误）。
        /// </summary>
        public static bool IsCallableBuiltin(string name) =>
            !string.IsNullOrEmpty(name) && _parens.Contains(name) && !_types.Contains(name);

        /// <summary>图标档位：kw / func / type；不是内建返回 null。</summary>
        public static string KindOf(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            if (_funcs.Contains(name)) return "func";
            if (_types.Contains(name)) return "type";
            if (_kw.Contains(name)) return "kw";
            return null;
        }

        /// <summary>给引擎用的内建名总表（关键字是否收由开关决定）。</summary>
        public static IEnumerable<string> Pool()
        {
            foreach (var n in _funcs) yield return n;
            foreach (var n in _types) yield return n;
            if (EnableKeywords)
                foreach (var n in _kw) yield return n;
        }

        public static int KeywordCount => _kw.Count;
        public static int FunctionCount => _funcs.Count + _types.Count;

        /// <summary>自检 / --dbtest 用：把内存镜像恢复成内置种子。</summary>
        public static void ResetToSeed() => ApplySeed();
    }
}
