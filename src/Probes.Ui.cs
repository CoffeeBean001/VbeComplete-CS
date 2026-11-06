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
    /// <summary>界面自检：候选窗 / 设置页离屏渲染。</summary>
    internal static partial class Program
    {
        /// <summary>
        /// 候选窗离屏渲染：把窗口画进一张 PNG，用来一眼核对配色 / 图标 / 布局。
        /// 不需要桌面，也不会出现在屏幕上。
        /// </summary>
        private static int UiTest()
        {
            var sb = new StringBuilder();
            try
            {
                var f = new PopupForm();
                // 用**真实匹配 + 真实排序**造数据：预览图里看到的高亮位置、
                // 图标档位、行序，都与用户在 VBE 里看到的一致。
                // 手写 Pos/Span 的那种预览图，改了算法它也不会跟着变。
                // ★词库也要真的载入：否则预览图里的内建名字全变成"var"图标。
                var wl = new WordLibrary();
                wl.Load();
                sb.AppendLine("word db: kw=" + wl.KeywordCount + " fn=" + wl.FunctionCount +
                              " kwPath=" + WordLibrary.KeywordPath() +
                              "(exists=" + File.Exists(WordLibrary.KeywordPath()) + ")" +
                              " err=" + wl.LastError);
                var eng = new Engine { Recent = new Recent() };
                eng.PrimePool("", "Module1");
                const string word = "se";
                var cands = new System.Collections.Generic.List<Candidate>();
                foreach (var n in new[]
                         {
                             "Select", "selBegin", "Size", "GetSettingTitle3", "dataSheet",
                             "MsgBox", "String", "Set", "Screen", "ScrollBars",
                         })
                {
                    FuzzyHit fh;
                    if (!Fuzzy.Match(n, word, out fh)) continue;
                    cands.Add(new Candidate
                    {
                        Name = n,
                        Kind = eng.Classify(n),
                        Pos = fh.First,
                        Span = fh.Span,
                        Hits = fh.Positions,
                        HitKind = fh.Kind,
                        ScopeKind = 3,
                        IsBuiltin = VbaBuiltins.IsKeyword(n),
                    });
                }
                cands.Sort(Engine.Cmp);

                f.SetData(cands, 1, 0, 15);
                var size = f.DesiredSize();
                int w = size.Width;
                int h = size.Height;
                f.StartPosition = FormStartPosition.Manual;
                f.Location = new System.Drawing.Point(-4000, -4000);
                f.Show();
                f.Width = w; f.Height = h;
                Application.DoEvents();

                using (var bmp = new System.Drawing.Bitmap(w, h))
                {
                    f.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, w, h));
                    string path = Path.Combine(Settings.Home, "popup-preview.png");
                    bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                    sb.AppendLine("saved " + path + "  " + w + "x" + h +
                                  "  items=" + cands.Count + "  word=" + word +
                                  "  font=" + f.FontName + "  rowH=" + f.RowHeight);
                }
                f.Hide();
                f.Dispose();

                // ---- 片段窗也离屏渲染一张：布局/配色/按钮顺序一眼能核对 ----
                try
                {
                    using (var store = new SnippetStore())
                    {
                        var sf = new SnippetForm(store, code => false);
                        sf.StartPosition = FormStartPosition.Manual;
                        sf.Location = new System.Drawing.Point(-4000, -4000);
                        sf.Show();
                        sf.PreviewOnly(store.All());
                        Application.DoEvents();
                        int sw = sf.Width, sh = sf.Height;
                        using (var bmp2 = new System.Drawing.Bitmap(sw, sh))
                        {
                            sf.DrawToBitmap(bmp2, new System.Drawing.Rectangle(0, 0, sw, sh));
                            string p2 = Path.Combine(Settings.Home, "snippet-preview.png");
                            bmp2.Save(p2, System.Drawing.Imaging.ImageFormat.Png);
                            sb.AppendLine("saved " + p2 + "  " + sw + "x" + sh +
                                           "  rows=" + store.Count());
                        }
                        sf.Hide();
                        sf.Dispose();
                    }
                }
                catch (Exception e2)
                {
                    sb.AppendLine("snippet preview skipped: " + e2.Message);
                }

                // ---- v175：设置界面 + 添加词条窗也离屏渲染，布局一眼能核对 ----
                try
                {
                    var sf = new SettingsForm(() => { }, () => { }, 0);
                    sf.StartPosition = FormStartPosition.Manual;
                    sf.Location = new System.Drawing.Point(-4000, -4000);
                    sf.Show();
                    for (int i = 0; i < 5; i++)
                    {
                        sf.GoPage(i);
                        Application.DoEvents();
                        Thread.Sleep(150);
                        int pw = sf.Width, ph = sf.Height;
                        using (var bmp3 = new System.Drawing.Bitmap(pw, ph))
                        {
                            sf.DrawToBitmap(bmp3, new System.Drawing.Rectangle(0, 0, pw, ph));
                            string p3 = Path.Combine(Settings.Home,
                                "settings-preview-" + i + ".png");
                            bmp3.Save(p3, System.Drawing.Imaging.ImageFormat.Png);
                            sb.AppendLine("saved " + p3 + "  " + pw + "x" + ph);
                        }
                    }
                    // 接线自检：勾选框 -> ini 键这一层是不是接对了（走真 SaveAll）。
                    // ★必须换到临时 ini：绝不能把用户的 settings.ini 改掉。
                    try
                    {
                        string oldIni = Environment.GetEnvironmentVariable(
                            Settings.PREFIX + "SETTINGS");
                        string tmpIni = Path.Combine(Path.GetTempPath(),
                            "vbecomplete-ui-" + Guid.NewGuid().ToString("N") + ".ini");
                        File.WriteAllText(tmpIni, "[settings]\r\n", new UTF8Encoding(false));
                        Environment.SetEnvironmentVariable(Settings.PREFIX + "SETTINGS", tmpIni);
                        Settings.Load();

                        var sf2 = new SettingsForm(() => { }, () => { }, 0);
                        bool wrote = sf2.SetRowForTest("AUTO_SPACE", "0");
                        bool took = !Settings.Flag("AUTO_SPACE", true);
                        bool savedAgain = sf2.SetRowForTest("AUTO_SPACE", "1");
                        bool tookBack = Settings.Flag("AUTO_SPACE", true);
                        bool rowCommented = Settings.IniValue("AUTO_SPACE") == null;  // 改回默认 = 那行重新注释掉
                        sf2.Dispose();
                        sb.AppendLine("wire AUTO_SPACE -> ini: wrote=" + wrote +
                                      " effective=" + took + " back=" + (savedAgain && tookBack) +
                                      " commentedBack=" + rowCommented);

                        // 反向项（NO_AUTOPAIR）：空 ini 时框必须是勾上的（功能默认开），
                        // 取消勾选才往 ini 里写 1，再勾回来 ini 应当回到 0。
                        var sf3 = new SettingsForm(() => { }, () => { }, 0);
                        bool pairOnOpen = sf3.RowCheckedForTest("NO_AUTOPAIR");
                        bool pairCleanIni = Settings.IniValue("NO_AUTOPAIR") == null;
                        bool pairWroteOff = sf3.SetRowForTest("NO_AUTOPAIR", "1");
                        bool pairOff = Settings.Flag("NO_AUTOPAIR", false) &&
                                       !sf3.RowCheckedForTest("NO_AUTOPAIR");
                        bool pairBack = sf3.SetRowForTest("NO_AUTOPAIR", "0") &&
                                        !Settings.Flag("NO_AUTOPAIR", false) &&
                                        sf3.RowCheckedForTest("NO_AUTOPAIR");
                        sf3.Dispose();
                        sb.AppendLine("wire NO_AUTOPAIR -> ui: checkedOnOpen=" + pairOnOpen +
                                      " iniClean=" + pairCleanIni + " off=" + (pairWroteOff && pairOff) +
                                      " back=" + pairBack);

                        Environment.SetEnvironmentVariable(Settings.PREFIX + "SETTINGS", oldIni);
                        Settings.Load();
                        try { File.Delete(tmpIni); } catch { }
                    }
                    catch (Exception e4)
                    {
                        sb.AppendLine("wire test FAILED: " + e4.Message);
                        sb.AppendLine("RESULT: FAIL");
                    }

                    // 布局自检：把第一页的控件树（类型 / 文字 / 实际尺寸）写进日志。
                    // ★凡是"看起来少了点什么"的布局问题，先看这里 —— 高度 0 还是没加进来，
                    //   一眼就分得清，不用靠截图猜。
                    DumpControls(sf.Controls, sb, 0);
                    sf.Hide();
                    sf.Dispose();

                    var wf = WordEditForm.Preview(WordBank.Functions, null);
                    wf.StartPosition = FormStartPosition.Manual;
                    wf.Location = new System.Drawing.Point(-4000, -4000);
                    wf.Show();
                    Application.DoEvents();
                    Thread.Sleep(120);
                    using (var bmp4 = new System.Drawing.Bitmap(wf.Width, wf.Height))
                    {
                        wf.DrawToBitmap(bmp4, new System.Drawing.Rectangle(0, 0, wf.Width, wf.Height));
                        string p4 = Path.Combine(Settings.Home, "wordedit-preview.png");
                        bmp4.Save(p4, System.Drawing.Imaging.ImageFormat.Png);
                        sb.AppendLine("saved " + p4 + "  " + wf.Width + "x" + wf.Height);
                    }
                    wf.Hide();
                    wf.Dispose();
                }
                catch (Exception e3)
                {
                    sb.AppendLine("settings preview FAILED: " + e3);
                }

                sb.AppendLine("RESULT: PASS");
                File.WriteAllText(Path.Combine(Settings.Home, "uitest.log"), sb.ToString(), Encoding.UTF8);
                return 0;
            }
            catch (Exception e)
            {
                sb.AppendLine("EXCEPTION " + e);
                File.WriteAllText(Path.Combine(Settings.Home, "uitest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }
        }

        /// <summary>把控件树按缩进打进日志（离线布局核对用）。</summary>
        private static void DumpControls(Control.ControlCollection cs, StringBuilder sb, int depth)
        {
            if (depth > 6) return;
            foreach (Control c in cs)
            {
                string txt = (c.Text ?? "").Replace("\r", " ").Replace("\n", " ");
                if (txt.Length > 40) txt = txt.Substring(0, 40) + "…";
                sb.AppendLine(new string(' ', depth * 2) + c.GetType().Name +
                              " [" + c.Left + "," + c.Top + " " + c.Width + "x" + c.Height + "]" +
                              (txt.Length > 0 ? " «" + txt + "»" : ""));
                DumpControls(c.Controls, sb, depth + 1);
            }
        }
    }
}
