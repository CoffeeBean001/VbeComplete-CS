using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace VbeComplete
{
    /// <summary>离线诊断：模块全文 dump、池子 dump。</summary>
    internal static partial class Program
    {
        /// <summary>
        /// 把当前代码窗模块的全文写到 moddump.txt（只读）。真机跑 --writetest / --hktest /
        /// --carettest 前 dump 一份、跑完再 dump 一份 diff —— 那几条用例都会"改完再删"，
        /// 这是唯一的核对手段。
        /// </summary>
        private static int DumpModule()
        {
            var be = new VbeBackend();
            string text = be.ModuleText();
            var snap = be.Snapshot();
            var sb = new StringBuilder();
            sb.AppendLine("module = " + (snap == null ? "?" : snap.Module) +
                          "  lines = " + be.LineCount() +
                          "  caret = " + (snap == null ? "?" : snap.Line + ":" + snap.CaretCol));
            sb.AppendLine("----");
            sb.Append(text ?? "(null)");
            File.WriteAllText(Path.Combine(Settings.Home, "moddump.txt"), sb.ToString(), Encoding.UTF8);
            return 0;
        }

        /// <summary>
        /// 诊断：把当前模块的解析结果 + 光标处的候选全写进 pooldump.txt。
        /// ★"某个名字该提示却没提示"时跑一次，就能分清是没进 FreeIdents、进了但作用域对不上、
        /// 还是被 Option Explicit 拦了。全程只读。
        /// </summary>
        private static int PoolDump(string onlyModule)
        {
            var sb = new StringBuilder();
            // 带模块名时：读那个模块的全文做"离线解析"（不跑 Trigger，只回答
            // "这个名字到底进没进 FreeIdents / 挂在哪个过程"）。
            if (!string.IsNullOrEmpty(onlyModule)) return PoolDumpNamed(onlyModule, sb);
            try
            {
                var be = new VbeBackend();
                string text = be.ModuleText();
                var snap = be.Snapshot();
                if (snap == null || text == null)
                {
                    sb.AppendLine("no code module（VBE 里没有打开任何代码窗）");
                    File.WriteAllText(Path.Combine(Settings.Home, "pooldump.txt"), sb.ToString(), Encoding.UTF8);
                    return 1;
                }

                var ms = Scope.Analyze(snap.Module, text);
                var proc = Parser.ProcOfLine(ms.Procs, snap.Line);
                string procName = proc == null ? "(null)" : proc.ScopeKey;

                sb.AppendLine("=== pool dump ===");
                sb.AppendLine("module   = " + snap.Module + "  lines = " + be.LineCount());
                sb.AppendLine("optionExplicit = " + ms.OptionExplicit);
                sb.AppendLine("caret    = L" + snap.Line + " C" + snap.CaretCol + "  line = [" + snap.Text + "]");
                sb.AppendLine("curProc  = " + procName);
                sb.AppendLine("procs    = " + string.Join(", ",
                    ms.Procs.Select(p => p.Name + ":" + p.StartLine + "-" + p.EndLine + "(" + p.ScopeKey + ")")));
                sb.AppendLine("decls    = " + string.Join(", ",
                    ms.Decls.Select(d => d.Name + "@" + (d.Proc ?? "-"))));
                sb.AppendLine("free     = " + string.Join(", ", ms.FreeIdents.Select(f =>
                {
                    HashSet<string> ps;
                    ms.FreeIdentProc.TryGetValue(f, out ps);
                    string where = ps == null || ps.Count == 0
                        ? "-"
                        : string.Join("/", ps.Select(x => x.Length == 0 ? "(模块级)" : x));
                    return f + "@" + where;
                })));
                sb.AppendLine("visible  = " + string.Join(", ", ms.Visible(procName).Keys.Take(80)));

                var eng = new Engine { Backend = be };
                bool shown = eng.Trigger(false, snap);
                sb.AppendLine("trigger  = " + shown + "  items = " + eng.Items.Count +
                              (eng.TruncatedFrom > 0 ? " (total " + eng.TruncatedFrom + ")" : ""));
                sb.AppendLine("cands    = " + string.Join(", ", eng.Items.Take(40).Select(c => c.Name)));
                // 判据逐条摊开：Trigger 静默返回 false 时，靠这几行定位是哪条挡的
                string lt = snap.Text ?? "";
                int cc = Rules.CharColFromDisp(lt, snap.CaretCol);
                int ws0;
                string w0 = Rules.WordBefore(lt, cc, out ws0);
                int ss = Rules.StatementStart(lt, Math.Max(0, cc - 1));
                int se = Math.Max(ss, Math.Min(cc - 1, lt.Length));
                sb.AppendLine("gates    = word=[" + w0 + "] ws=" + ws0 + " cc=" + cc +
                              " seg=[" + lt.Substring(ss, se - ss) + "]" +
                              " dead=" + Rules.InDeadZone(lt, cc) +
                              " declPos=" + Rules.InDeclarationPosition(lt, cc) +
                              " afterType=" + Rules.AfterTypeKeyword(lt, ws0) +
                              " dotCol=" + Rules.MemberDotCol(lt, cc) +
                              " pool=" + eng.PoolCount);
            }
            catch (Exception e)
            {
                sb.AppendLine("EXCEPTION " + e);
            }
            File.WriteAllText(Path.Combine(Settings.Home, "pooldump.txt"), sb.ToString(), Encoding.UTF8);
            return 0;
        }

        /// <summary>按模块名离线解析（不需要那个代码窗是活动的）。</summary>
        private static int PoolDumpNamed(string name, StringBuilder sb)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) { sb.AppendLine("no VBE"); return 1; }
                dynamic comps = vbe.ActiveVBProject.VBComponents;
                int n = Convert.ToInt32(comps.Count);
                for (int i = 1; i <= n; i++)
                {
                    dynamic c = comps.Item(i);
                    string nm = Convert.ToString(c.Name) ?? "";
                    if (!string.Equals(nm, name, StringComparison.OrdinalIgnoreCase)) continue;
                    dynamic cm = c.CodeModule;
                    int cnt = Convert.ToInt32(cm.CountOfLines);
                    string text = cnt > 0 ? Convert.ToString(cm.Lines(1, cnt)) : "";
                    var ms = Scope.Analyze(nm, text ?? "");
                    sb.AppendLine("=== pool dump (module " + nm + ", lines " + cnt + ") ===");
                    sb.AppendLine("optionExplicit = " + ms.OptionExplicit);
                    sb.AppendLine("procs    = " + string.Join(", ",
                        ms.Procs.Select(p => p.Name + ":" + p.StartLine + "-" + p.EndLine + "(" + p.ScopeKey + ")")));
                    sb.AppendLine("decls    = " + string.Join(", ",
                        ms.Decls.Select(d => d.Name + "@" + (d.Proc ?? "-"))));
                    sb.AppendLine("free     = " + string.Join(", ", ms.FreeIdents.Select(f =>
                    {
                        HashSet<string> ps;
                        ms.FreeIdentProc.TryGetValue(f, out ps);
                        string where = ps == null || ps.Count == 0
                            ? "-"
                            : string.Join("/", ps.Select(x => x.Length == 0 ? "(模块级)" : x));
                        return f + "@" + where;
                    })));
                    File.WriteAllText(Path.Combine(Settings.Home, "pooldump.txt"), sb.ToString(), Encoding.UTF8);
                    return 0;
                }
                sb.AppendLine("module not found: " + name);
            }
            catch (Exception e) { sb.AppendLine("EXCEPTION " + e); }
            File.WriteAllText(Path.Combine(Settings.Home, "pooldump.txt"), sb.ToString(), Encoding.UTF8);
            return 1;
        }

        /// <summary>
        /// 写入链路验证：在模块末尾追加一行 -> 用补全改它 -> 删掉，最后逐字节比对原文。
        /// ★只在用户显式跑 `--writetest` 时动，正常运行一次都不会碰用户代码。
        /// </summary>
        private static int WriteTest()
        {
            var sb = new StringBuilder();
            var be = new VbeBackend();
            string orig = be.ModuleText();
            int count = be.LineCount();
            bool ok = true;

            sb.AppendLine("=== write test ===");
            sb.AppendLine("module lines = " + count + "  text len = " + (orig == null ? -1 : orig.Length));
            if (count <= 0)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "writetest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }

            try
            {
                const string probe = "    Dim abc";
                if (!be.InsertLines(count + 1, probe))
                {
                    sb.AppendLine("InsertLines FAILED");
                    ok = false;
                }
                else
                {
                    string back = be.LineAt(count + 1);
                    sb.AppendLine("insert readback = [" + back + "]");
                    bool same = back == probe;
                    if (!same) ok = false;
                    sb.AppendLine("insert roundtrip = " + same);

                    bool w = be.ApplyCompletion(count + 1, 9, 12, "abcd", false, false, null);
                    string after = be.LineAt(count + 1);
                    sb.AppendLine("apply ok=" + w + "  result=[" + after + "]");
                    sb.AppendLine("apply result correct = " + (after == "    Dim abcd"));
                    if (after != "    Dim abcd") ok = false;

                    be.SetCaret(count + 1, 13);
                    var s2 = be.Snapshot();
                    string got = s2 == null ? "null" : s2.Line + ":" + s2.CaretCol;
                    sb.AppendLine("caret after SetCaret(13) = " + got + " (want 13)");
                    if (s2 == null || s2.CaretCol != 13) ok = false;
                }
            }
            finally
            {
                // 无论成败都要把测试行删干净（最多重试 3 次）
                for (int i = 0; i < 3; i++)
                {
                    int c2 = be.LineCount();
                    if (c2 <= count) break;
                    be.DeleteLines(count + 1, c2 - count);
                }
            }

            string now = be.ModuleText();
            bool restored = now == orig;
            sb.AppendLine("restored = " + restored);
            if (!restored)
            {
                ok = false;
                sb.AppendLine("!!! 未能逐字节还原，原文如下（请手工检查模块末尾） !!!");
                sb.AppendLine(orig);
                sb.AppendLine("!!! 现在的全文 !!!");
                sb.AppendLine(now);
            }
            sb.AppendLine(ok && restored ? "RESULT: PASS" : "RESULT: FAIL");

            File.WriteAllText(Path.Combine(Settings.Home, "writetest.log"), sb.ToString(), Encoding.UTF8);
            return ok && restored ? 0 : 1;
        }
        /// </summary>
        /// <summary>日志用：安全截一段字符串（越界不抛，换行写成 \n）。</summary>
        private static string SafeSlice(string s, int from, int len)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (from < 0) from = 0;
            if (from >= s.Length) return "";
            string part = s.Substring(from, Math.Min(len, s.Length - from));
            return part.Replace("\r", "\\r").Replace("\n", "\\n");
        }
    }
}
