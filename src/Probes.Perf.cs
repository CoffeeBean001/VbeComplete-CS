// 性能探针：候选弹窗这条路，每一段各花多少毫秒。
//
// ★为什么要有它：用户说"提示词列表弹出速度有点慢"，而这句话背后至少有四段开销
//   （轮询间隔 / 读快照读全文 / 攒池子解析 / 匹配排序画窗），凭感觉猜一定猜错。
//   这个探针把四段分别量出来，改前改后各跑一次，谁的功劳一目了然。
// ★两半：CPU 那半用合成工程（不需要 Excel，随时能跑）；真机那半读当前打开的模块，
//   没有代码窗就 SKIP —— "没测到"必须自己说出来，不许装绿。

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace VbeComplete
{
    internal static partial class Program
    {
        /// <summary>`--perfprobe`：候选弹窗各段耗时。只读，不改模块、不改设置。</summary>
        private static int PerfProbe()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== perf probe（候选弹窗各段耗时） ===");
            sb.AppendLine("settings: MAX=" + Environment.ProcessorCount + " core  词库 Pool=" +
                          VbaBuiltins.Pool().Count() + " 条");
            sb.AppendLine();

            CpuBench(sb);
            sb.AppendLine();
            LiveBench(sb);
            sb.AppendLine();
            E2eBench(sb);

            string path = Path.Combine(Settings.Home, "perfprobe.log");
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
            Console.WriteLine(sb.ToString());
            Console.WriteLine("（同时写到 " + path + "）");
            // 真机那两段没跑到时用 3 表示 SKIP —— 跟别的探针一个口径
            return sb.ToString().Contains("SKIP") ? 3 : 0;
        }

        // ------------------------------------------------------------------
        // CPU：合成工程，不碰 COM
        // ------------------------------------------------------------------

        /// <summary>造一个"像真工程"的模块：声明、过程、隐式变量都有。</summary>
        private static string BenchModule(int seed, int lines)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Option Explicit");
            sb.AppendLine();
            for (int i = 0; i < 12; i++)
                sb.AppendLine("Private g" + seed + "_" + i + " As Long");
            sb.AppendLine("Public Type T" + seed);
            sb.AppendLine("    a As Long");
            sb.AppendLine("    b As String");
            sb.AppendLine("End Type");
            sb.AppendLine();
            int proc = 0;
            while (true)
            {
                sb.AppendLine("Public Function Fn" + seed + "_" + proc + "(ByVal n As Long) As Long");
                for (int i = 0; i < 10; i++)
                {
                    sb.AppendLine("    Dim v" + proc + "_" + i + " As Long");
                    sb.AppendLine("    v" + proc + "_" + i + " = n * " + i + " + g" + seed + "_" + i);
                }
                sb.AppendLine("    Fn" + seed + "_" + proc + " = v" + proc + "_0");
                sb.AppendLine("End Function");
                sb.AppendLine();
                proc++;
                // 行数是估算的：只要到量级就行，精确到哪一行没意义
                if (proc * 28 >= lines) break;
            }
            return sb.ToString();
        }

        private static void CpuBench(StringBuilder sb)
        {
            const int Modules = 8;
            const int LinesPerModule = 200;
            var texts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < Modules; i++)
                texts["模块" + i] = BenchModule(i, LinesPerModule);

            string cur = "模块0";
            string baseText = texts[cur];
            int baseLines = baseText.Replace("\r\n", "\n").Split('\n').Length;

            var be = new CachedBackend(texts, cur);
            var eng = new Engine { Backend = be, Recent = null };
            // 同一份工程跑两档：NoPoolCache=1 就是 v189 的老行为（每键重建池子），
            // =0 是 v190（同一个词里连着打只解析一次）。两档同一次运行里量，才可比。
            var engOld = new Engine { Backend = be, Recent = null, NoPoolCache = true };

            sb.AppendLine("---- CPU（合成工程 " + Modules + " 个模块 × ~" + LinesPerModule +
                          " 行，当前模块 " + baseLines + " 行）----");

            // 逐字符"打字"：每敲一个字母就是一次完整的候选计算，这正是用户感知的那一跳
            string[] typed = { "m", "ms", "msg", "msgb", "msgbo" };
            var tTotal = new List<double>();
            var tOld = new List<double>();
            var tNewWord = new List<double>();     // 每敲都换一个"新词"（池子必然重建）
            var tAnalyze = new List<double>();

            for (int round = 0; round < 8; round++)
            {
                foreach (var q in typed)
                {
                    string text = baseText + "\r\n    y = " + q;
                    string line = "    y = " + q;
                    var snap = new Snapshot
                    {
                        Line = baseLines + 1,
                        Text = line,
                        Module = cur,
                        CaretCol = line.Length + 1,     // 显示列（纯 ASCII，= 字符列）
                    };

                    var sw = Stopwatch.StartNew();
                    eng.Trigger(false, snap);
                    sw.Stop();
                    tTotal.Add(sw.Elapsed.TotalMilliseconds);

                    engOld.InvalidatePool();
                    sw.Restart();
                    engOld.Trigger(false, snap);
                    sw.Stop();
                    tOld.Add(sw.Elapsed.TotalMilliseconds);

                    // 只解析当前模块（池子里最贵的就是这一步）
                    sw.Restart();
                    Scope.Analyze(cur, text);
                    sw.Stop();
                    tAnalyze.Add(sw.Elapsed.TotalMilliseconds);
                }

                // 池子真得重建的那一档（换了词 / 别处改了代码）—— 用显式作废来量
                foreach (var q in typed)
                {
                    string line = "    y = " + q;
                    var snap2 = new Snapshot
                    {
                        Line = baseLines + 1,
                        Text = line,
                        Module = cur,
                        CaretCol = line.Length + 1,
                    };
                    eng.InvalidatePool();
                    var sw = Stopwatch.StartNew();
                    eng.Trigger(false, snap2);
                    sw.Stop();
                    tNewWord.Add(sw.Elapsed.TotalMilliseconds);
                }
            }

            var pool = VbaBuiltins.Pool().ToList();
            var tMatch = new List<double>();
            for (int i = 0; i < 8; i++)
            {
                var sw = Stopwatch.StartNew();
                int hit = 0;
                foreach (var n in pool)
                    if (Fuzzy.Match(n, "msg", out FuzzyHit fh)) hit++;
                sw.Stop();
                tMatch.Add(sw.Elapsed.TotalMilliseconds);
            }

            sb.AppendLine("  同一个词里连打（v190）        " + Stat(tTotal) + "  池子 " + eng.PoolCount + " 条");
            sb.AppendLine("  同一个词里连打（v189 老键）  " + Stat(tOld) + "  ← NO_POOL_CACHE=1");
            sb.AppendLine("  池子必须重建的那一档         " + Stat(tNewWord));
            sb.AppendLine("  其中 Scope.Analyze(当前模块)  " + Stat(tAnalyze));
            sb.AppendLine("  纯 Fuzzy.Match 扫一遍内建库(" + pool.Count + " 条)  " + Stat(tMatch));

            // ---- 大模块：真实工程里一个模块动辄上千行，而 Scope.Analyze 是**按行全量**
            //      重解析的。每开始一个新词就要重解析一遍 —— 这条是"大工程里打字发涩"的
            //      唯一嫌疑，必须单独量。
            sb.AppendLine();
            sb.AppendLine("---- CPU（大模块：1 个 " + BigLines + " 行的当前模块）----");
            var big = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["大模块"] = BenchModule(99, BigLines),
            };
            var bigBe = new CachedBackend(big, "大模块");
            var bigEng = new Engine { Backend = bigBe, Recent = null };
            string bigText = big["大模块"];
            int bigLines = bigText.Replace("\r\n", "\n").Split('\n').Length;
            var tBigAnalyze = new List<double>();
            var tBigNewWord = new List<double>();
            for (int i = 0; i < 10; i++)
            {
                var sw2 = Stopwatch.StartNew();
                Scope.Analyze("大模块", bigText);
                sw2.Stop();
                tBigAnalyze.Add(sw2.Elapsed.TotalMilliseconds);

                // 每敲一个字母都是**新词**（词起点变了）⇒ 池子必须重建
                string ln = "    y" + (char)('a' + i) + "z = 1";
                var sn = new Snapshot
                {
                    Line = bigLines + 1, Text = ln, Module = "大模块", CaretCol = ln.Length + 1,
                };
                bigEng.InvalidatePool();
                sw2.Restart();
                bigEng.Trigger(false, sn);
                sw2.Stop();
                tBigNewWord.Add(sw2.Elapsed.TotalMilliseconds);
            }
            sb.AppendLine("  每开始一个新词（整条 CPU）    " + Stat(tBigNewWord) + "  池子 " + bigEng.PoolCount + " 条");
            sb.AppendLine("  其中 Scope.Analyze(" + bigLines + " 行)      " + Stat(tBigAnalyze));

            // 再切一刀：Analyze 里"去注释/去字符串"（逐行扫描）和"提取过程/类型"各占多少
            var bigLinesArr = Parser.SplitLines(bigText);
            var tCodeOnly = new List<double>();
            var tProcs = new List<double>();
            for (int i = 0; i < 10; i++)
            {
                var sw3 = Stopwatch.StartNew();
                for (int j = 0; j < bigLinesArr.Length; j++) Parser.CodeOnly(bigLinesArr[j]);
                sw3.Stop();
                tCodeOnly.Add(sw3.Elapsed.TotalMilliseconds);

                sw3.Restart();
                Parser.ExtractProcedures(bigText);
                sw3.Stop();
                tProcs.Add(sw3.Elapsed.TotalMilliseconds);
            }
            sb.AppendLine("  其中 逐行 CodeOnly          " + Stat(tCodeOnly));
            sb.AppendLine("  其中 ExtractProcedures      " + Stat(tProcs));
            sb.AppendLine("  Analyze 分段：切行 " + Scope.ProfSplitMs + " / 声明 " + Scope.ProfDeclMs +
                          " / 用法 " + Scope.ProfUsageMs + " / 写入 " + Scope.ProfWrittenMs +
                          " / 裸名 " + Scope.ProfFreeMs + " ms");
        }

        private const int BigLines = 2000;

        /// <summary>中位数 / p95 / 均值，一行写完。</summary>
        private static string Stat(List<double> xs)
        {
            if (xs == null || xs.Count == 0) return "n/a";
            var a = xs.OrderBy(x => x).ToList();
            double med = a[a.Count / 2];
            double p95 = a[Math.Min(a.Count - 1, (int)(a.Count * 0.95))];
            return string.Format("中位 {0,6:F3} ms   p95 {1,6:F3} ms   均 {2,6:F3} ms  (n={3})",
                med, p95, a.Average(), a.Count);
        }

        // ------------------------------------------------------------------
        // 真机：读快照 / 读全文 / 整条 Trigger
        // ------------------------------------------------------------------

        private static void LiveBench(StringBuilder sb)
        {
            sb.AppendLine("---- 真机（当前打开的代码窗）----");
            var be = new VbeBackend();
            Snapshot snap0 = be.Snapshot();
            string text0 = be.ModuleText();
            if (snap0 == null || string.IsNullOrEmpty(text0))
            {
                sb.AppendLine("SKIP：VBE 里没有打开的代码窗，真机这一段一个字都没测到。");
                return;
            }
            string cur = snap0.Module ?? "";
            int lines = text0.Replace("\r\n", "\n").Split('\n').Length;
            sb.AppendLine("  模块 = " + cur + "  行 = " + lines + "  长度 = " + text0.Length);

            // ① 读快照（禁用缓存：打字时每轮都会作废，等于每轮都真读）
            var t = new List<double>();
            for (int i = 0; i < 30; i++)
            {
                be.InvalidateSnapshot();
                var sw = Stopwatch.StartNew();
                be.Snapshot();
                sw.Stop();
                t.Add(sw.Elapsed.TotalMilliseconds);
            }
            sb.AppendLine("  Snapshot()（不过缓存）         " + Stat(t));

            // ② 读模块全文（冷读）
            t.Clear();
            for (int i = 0; i < 15; i++)
            {
                be.ForceReloadModule();
                var sw = Stopwatch.StartNew();
                be.ModuleText();
                sw.Stop();
                t.Add(sw.Elapsed.TotalMilliseconds);
            }
            sb.AppendLine("  ModuleText()（冷读全文）       " + Stat(t));

            // ③ 全工程作用域（命中缓存 / 强制重算）
            t.Clear();
            for (int i = 0; i < 20; i++)
            {
                var sw = Stopwatch.StartNew();
                be.ProjectScopes(out _);
                sw.Stop();
                t.Add(sw.Elapsed.TotalMilliseconds);
            }
            sb.AppendLine("  ProjectScopes()（走缓存）      " + Stat(t));

            // ④ 真机整条 Trigger：池子作废 = 每键最坏情况；池子在 = 只匹配排序
            var eng = new Engine { Backend = be, Recent = null };
            var snap = be.Snapshot();
            var tCold = new List<double>();
            var tWarm = new List<double>();
            for (int i = 0; i < 15; i++)
            {
                eng.InvalidatePool();
                var sw = Stopwatch.StartNew();
                eng.Trigger(false, snap);
                sw.Stop();
                tCold.Add(sw.Elapsed.TotalMilliseconds);

                sw.Restart();
                eng.Trigger(false, snap);
                sw.Stop();
                tWarm.Add(sw.Elapsed.TotalMilliseconds);
            }
            sb.AppendLine("  Trigger()（池子重建整条）      " + Stat(tCold) + "  池子 " + eng.PoolCount + " 条");
            sb.AppendLine("  Trigger()（池子命中）          " + Stat(tWarm));
            eng.Hide();
        }

        // ------------------------------------------------------------------
        // 端到端：敲一个字符 → 轮到候选计算跑完，中间隔多久
        // ------------------------------------------------------------------

        // 打字前切到的输入法档位（只为收尾还原；perfprobe 单跑，不做并发保护）
        private static int _imeMode0 = -1;
        private static bool _imeUsedShift;
        private static long _e2ePollBase;      // 每一轮采样前的 PollCount（算"这一键跑了几轮"）

        /// <summary>
        /// 用户感知的就是这一段：手指按下 → 候选窗在屏幕上。
        /// 钩子只标"刚按过键"，真正算候选的是轮询 —— 所以这段延迟主要是
        /// "轮询多久才轮到这一轮"，正是本次优化要动的地方。
        /// </summary>
        private static void E2eBench(StringBuilder sb)
        {
            sb.AppendLine("---- 端到端（真按键：敲一个字符 → 候选计算跑完）----");
            var app = new App();
            var be = new VbeBackend();
            var gate = new object();
            var samples = new List<double>();
            var wake = new List<double>();     // 按键 → 轮询开始（纯"等"）
            var polls = new List<double>();    // 这一次按键一共跑了几轮轮询
            var traces = new List<string>();   // 每次唤醒的补跳轨迹（p=轮询耗时 d=等待）
            var trigCost = new List<double>(); // 其中 Trigger 自己（匹配/排序/画窗）花了多少
            var prof = new List<string>();     // Trigger 内部分段：读文/池/匹配/排序/画窗
            string note = "";

            int baseCount = be.LineCount();
            Snapshot s0 = be.Snapshot();
            if (baseCount < 0 || s0 == null)
            {
                sb.AppendLine("SKIP：没有代码窗，端到端这一段没测到。");
                return;
            }
            string orig = be.ModuleText();
            int at = baseCount + 1;
            // ★必须真启动：钩子和轮询都在 Start 里装上 —— 不启的话既收不到键，也不会有
            //   "轮询多久才轮到这一轮"，端到端那条就永远采不到（第一版漏了这个，白跑一轮）。
            be.FocusVbe();
            app.Start();
            // 等应用自己的预热走完（候选窗 1.2s、全工程作用域同一次）—— 用户真实打字时
            // 那一步早付掉了，这里不等的话量到的是"冷启动"，没有代表性。
            Thread.Sleep(1500);

            var driver = new Thread(() =>
            {
                try
                {
                    be.FocusVbe();
                    bool focused = false;
                    for (int i = 0; i < 40 && !focused; i++)
                    {
                        app.RunOnUi(() =>
                        {
                            Native.ForceForeground(Native.FindWindow("wndclass_desked_gsk", null));
                            be.SetCaret(at + 1, 5);
                            be.InvalidateSnapshot();
                        });
                        Thread.Sleep(250);
                        focused = app.CodePaneFocused;
                    }
                    if (!focused)
                    {
                        string fg = Native.ClassNameOf(Native.GetForegroundWindow());
                        lock (gate) note = "SKIP：焦点抢不到代码窗（前台窗口类 = " + fg +
                                           "，VBE 主窗 = " + Native.FindWindow("wndclass_desked_gsk", null) + "）";
                        return;
                    }

                    // ★必须先把输入法切到英文档：中文档下敲 `B` 只会开始组字（文档一个字都不变），
                    //   量到的就不是"弹窗慢"，而是"根本没敲进去"。跑完照原样还回去。
                    int imeMode0 = -1;
                    bool imeUsedShift = false;
                    bool alnum = PrepareImeForTyping(
                        s => { lock (gate) sb.AppendLine("      " + s); },
                        out imeMode0, out imeUsedShift);
                    _imeMode0 = imeMode0;
                    _imeUsedShift = imeUsedShift;
                    if (!alnum)
                    {
                        lock (gate) note = "SKIP：输入法切不到英文档，敲不进去（这一轮没测到）";
                        return;
                    }

                    for (int k = 0; k < 10; k++)
                    {
                        // 摆一行半截词（不带最后一个字符），让"敲最后一下"才有东西可弹
                        RebuildScaffold(app, be, at, "    y = Msg", "    Dim resProbe As Long");
                        Thread.Sleep(320);
                        app.RunOnUi(() =>
                        {
                            app.ProbeForgetLastTrigger();
                            _e2ePollBase = app.PollCount;
                        });
                        SendKeys(app, (int)'B');
                        // 采样只读字段，不投递 UI 动作 —— 投递一次就要 1ms 上下，会糊掉读数
                        var sw = Stopwatch.StartNew();
                        long got = -1;
                        while (sw.ElapsedMilliseconds < 1200)
                        {
                            got = app.LastKeyToTriggerMs;
                            if (got >= 0) break;
                            Thread.SpinWait(60);
                        }
                        lock (gate)
                        {
                            if (got >= 0)
                            {
                                samples.Add(got);
                                long toPoll = app.LastKeyToPollMs;
                                if (toPoll >= 0) wake.Add(toPoll);
                                polls.Add((double)(app.PollCount - _e2ePollBase));
                                long cost = app.LastTriggerCostMs;
                                if (cost >= 0) trigCost.Add(cost);
                                var eg = app.Eng;
                                if (eg != null && eg.ProfShowMs >= 0)
                                    prof.Add(eg.ProfTextMs + "/" + eg.ProfPoolMs + "/" +
                                             eg.ProfMatchMs + "/" + eg.ProfSortMs + "/" +
                                             eg.ProfShowMs);
                                // 前三次把"等多久 / 每轮轮询花多久"的轨迹抄下来
                                if (traces.Count < 3 && app.BurstTrace != null)
                                    traces.Add(app.BurstTrace);
                            }
                            else note = "有一次 1200ms 内没等到候选计算（宿主被挤？）";
                        }
                        Thread.Sleep(200);
                    }
                }
                catch (Exception e) { lock (gate) note = "EXCEPTION " + e.Message; }
                finally
                {
                    // 输入法原样还回去（与 declprobe 同一套收尾）
                    Native.ImeRestore(Native.FocusedVbeWindow(), _imeMode0);
                    if (_imeUsedShift && !Native.ImeNativeOn(Native.FocusedVbeWindow()))
                        Native.SendVk(Native.VK_SHIFT);
                    app.RequestExit();
                }
            })
            { IsBackground = true, Name = "perfprobe" };

            driver.Start();
            try
            {
                Application.Run(app);
            }
            finally
            {
                driver.Join(4000);
                try
                {
                    // 脚手架必删：这段是往用户真模块里写行，留着就是脏数据
                    int tailCnt = be.LineCount();
                    if (tailCnt >= at) be.DeleteLines(at, tailCnt - at + 1);
                    be.ForceReloadModule();
                    CleanupScratch(be);
                    be.ForceReloadModule();
                    bool restored = string.Equals(be.ModuleText(), orig, StringComparison.Ordinal);
                    lock (gate)
                    {
                        if (samples.Count > 0)
                        {
                            sb.AppendLine("  按键 → 候选算完（Trigger 跑完）  " + Stat(samples));
                            if (wake.Count > 0)
                                sb.AppendLine("  其中：按键 → 轮询开始（纯等）   " + Stat(wake));
                            if (polls.Count > 0)
                                sb.AppendLine("  每个键平均跑几轮轮询           " +
                                    polls.Average().ToString("F2") + " 轮" +
                                    "（上限 " + polls.Max().ToString("F0") + "）");
                            if (trigCost.Count > 0)
                                sb.AppendLine("  其中：Trigger 自己（含画窗）    " + Stat(trigCost));
                            foreach (var tr in traces)
                                sb.AppendLine("  轨迹：" + tr);
                            foreach (var pr in prof)
                                sb.AppendLine("  分段(读文/池/匹配/排序/画窗 ms)：" + pr);
                        }
                        else
                            sb.AppendLine("SKIP：一次采样都没拿到。");
                        if (note.Length > 0) sb.AppendLine("  注：" + note);
                        sb.AppendLine("  cleanup restored = " + restored);
                    }
                }
                catch (Exception e) { lock (gate) sb.AppendLine("cleanup 异常：" + e.Message); }
            }
        }

        /// <summary>
        /// 合成后端，但**照真机那样缓存解析结果**：真机 ProjectScopes 2 秒才重算一次，
        /// 每次调用只做"取缓存"。这里如果每调一次就把所有模块重新解析，量到的就是
        /// 探针自己的开销，而不是产品的。
        /// </summary>
        private sealed class CachedBackend : IVbeSource
        {
            private readonly List<ModuleScope> _list = new List<ModuleScope>();
            private readonly Dictionary<string, string> _texts;
            private readonly string _cur;

            public CachedBackend(Dictionary<string, string> texts, string cur)
            {
                _texts = texts;
                _cur = cur;
                foreach (var kv in texts)
                {
                    var ms = Scope.Analyze(kv.Key, kv.Value);
                    ms.ComponentType = 1;                     // 都当标准模块
                    _list.Add(ms);
                }
            }

            public Snapshot Snapshot(bool verify = true) => null;
            public string ModuleText() => _texts.TryGetValue(_cur, out var t) ? t : "";
            public List<string> ComponentNames() => new List<string>(_texts.Keys);
            public List<ModuleScope> ProjectScopes(out int signature)
            {
                signature = 1;
                return _list;
            }
            public bool ApplyCompletion(int lineNo, int wordStartCol, int endCol,
                string completion, bool wantParens, bool wantSpace, string enclosingProc) => false;
            public string EnclosingProcName() => null;
        }
    }
}
