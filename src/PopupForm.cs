// 候选窗：无边框、置顶、【永不抢焦点】、全部自绘。
//
// ★四条不能动的规矩：
//   1. 绝不抢焦点（WS_EX_NOACTIVATE + ShowWithoutActivation）—— 抢一下焦点，
//      用户在 VBE 里的输入就断了，整个工具等于废掉；
//   2. 全部自绘（一个 OnPaint 画完）—— 每行一个 Control 的话，滚一下要重排
//      几十个句柄，那才是"卡"的来源；
//   3. 拿不到光标矩形就【不显示】，绝不猜一个位置糊上去；
//   4. 宽度【固定】为 30 个字符 —— 随内容伸缩会让窗口每敲一个字都横向抖一下，
//      眼睛一直跟着跳（IDEA / VBE 自己的成员列表都是固定宽度）。
//
// ★v132 对齐 Python 版 ui.py 的几处观感：
//   * 逐字符高亮（不是"一段区间"）：打 `ds` 命中 dataSheet 的 D 与 S 两个字母；
//   * 鼠标悬停行有浅蓝底，跟选中的深蓝区分开；
//   * 详情行（完整名字）只在名字被省略时出现 —— 否则白占一行高度；
//   * 一屏最多 15 行，候选不足时窗口跟着变矮；
//   * 候选窗若与 VBE 自己的参数信息窗重叠，整体下移到它下方。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace VbeComplete
{
    internal sealed class PopupForm : Form
    {
        // ---- 配色（与 python 版 ui.py 顶部常量逐项对应） ----
        private static readonly Color CBg = ColorTranslator.FromHtml("#FFFFFF");
        private static readonly Color CFg = ColorTranslator.FromHtml("#202020");
        private static readonly Color CHit = ColorTranslator.FromHtml("#C00000");
        private static readonly Color CSelBg = ColorTranslator.FromHtml("#2A5CAA");
        private static readonly Color CSelFg = ColorTranslator.FromHtml("#FFFFFF");
        private static readonly Color CSelHit = ColorTranslator.FromHtml("#FFD24A");
        private static readonly Color CHoverBg = ColorTranslator.FromHtml("#EAF1FB");
        private static readonly Color CDetailBg = ColorTranslator.FromHtml("#F7F7F7");
        private static readonly Color CDetailFg = ColorTranslator.FromHtml("#333333");
        private static readonly Color CDetailSep = ColorTranslator.FromHtml("#E0E0E0");
        private static readonly Color CTrack = ColorTranslator.FromHtml("#F0F0F0");
        private static readonly Color CThumb = ColorTranslator.FromHtml("#C8C8C8");
        private static readonly Color CBorder = ColorTranslator.FromHtml("#B8B8B8");

        private static readonly Dictionary<string, Color> IconColor =
            new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase)
            {
                { "func", ColorTranslator.FromHtml("#1F8A3B") },
                { "prop", ColorTranslator.FromHtml("#7A3FA8") },
                { "var",  ColorTranslator.FromHtml("#2A5CAA") },
                { "type", ColorTranslator.FromHtml("#B26A00") },
                { "kw",   ColorTranslator.FromHtml("#8A8A8A") },
            };

        private static readonly Dictionary<string, Color> IconColorOn =
            new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase)
            {
                { "func", ColorTranslator.FromHtml("#6FE08C") },
                { "prop", ColorTranslator.FromHtml("#CBA6F7") },
                { "var",  ColorTranslator.FromHtml("#9CC2F5") },
                { "type", ColorTranslator.FromHtml("#FFC46B") },
                { "kw",   ColorTranslator.FromHtml("#D8D8D8") },
            };

        private const int PadX = 6;
        private const int IconGap = 4;
        private const int ScrollW = 8;
        private const int NameChars = 30;        // 列表宽度固定容纳这么多字符
        private const string Ellipsis = "...";
        private const int MaxWidth = 640;
        private const int PopupGap = 2;          // 避让 VBE 提示窗时的留白
        private const int PopupTryMax = 4;

        private IReadOnlyList<Candidate> _items = Array.Empty<Candidate>();
        private int _sel;
        private int _top;
        private int _rows = 8;
        private int _rowH = 20;
        private int _iconW = 12;
        private int _hover = -1;
        private int _lastX = int.MinValue, _lastY = int.MinValue, _lastW, _lastH;

        private bool _dragScroll;
        private int _dragOffset;

        private readonly Font _font;
        private int _charW;          // 一个字符的像素宽（按 "x"*NameChars 量）

        public event Action<int> RowClicked;      // 单击某行（index 为绝对下标）
        public event Action<int> Wheel;           // 滚轮行数
        public event Action<int> ScrollTo;        // 滚动条拖动：目标 top

        public PopupForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = CBg;
            DoubleBuffered = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);

            _font = BuildFont(out float pt);
            _rowH = Math.Max(16, _font.Height + 6);
            _iconW = Math.Max(9, Math.Min(16, (int)Math.Round(_rowH * 0.62)));
            _charW = Math.Max(4, MeasureText(new string('x', NameChars)) / NameChars);

            TabStop = false;
        }

        /// <summary>
        /// 跟 VBE 编辑器用同一个字体、**小一号**（Python 版 FONT_DELTA = -1）——
        /// 候选窗的字跟编辑器一模一样大反而喧宾夺主。注册表读不到就退回 Consolas 10。
        /// </summary>
        private static Font BuildFont(out float pt)
        {
            double want = Settings.Double("POPUP_FONT_SIZE", 0);
            string face = Settings.Raw("POPUP_FONT");
            try
            {
                if (string.IsNullOrEmpty(face) || want <= 0)
                {
                    using (var k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VBA\6.0\Common"))
                    {
                        if (k != null)
                        {
                            if (string.IsNullOrEmpty(face))
                                face = k.GetValue("FontFace") as string;
                            if (want <= 0)
                            {
                                object h = k.GetValue("FontHeight");
                                if (h is int i && i > 0) want = i;
                            }
                        }
                    }
                }
            }
            catch { }

            if (want <= 0) want = 10;
            else want -= 1;                  // 比编辑器小一号
            if (want < 8) want = 8;
            if (want > 24) want = 24;
            pt = (float)want;
            if (string.IsNullOrEmpty(face)) face = "Consolas";

            try
            {
                var f = new Font(face, pt, FontStyle.Regular);
                if (string.Equals(f.Name, face, StringComparison.OrdinalIgnoreCase)) return f;
                f.Dispose();
            }
            catch { }
            return new Font("Consolas", pt, FontStyle.Regular);
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000;   // WS_EX_NOACTIVATE：点了也不抢焦点
                cp.ExStyle |= 0x00000080;   // WS_EX_TOOLWINDOW：不出现在 Alt-Tab
                return cp;
            }
        }

        public IReadOnlyList<Candidate> Items => _items;
        public int RowHeight => _rowH;

        public void SetData(IReadOnlyList<Candidate> items, int sel, int top, int rows)
        {
            _items = items ?? Array.Empty<Candidate>();
            _sel = sel;
            _top = top;
            _rows = Math.Max(1, rows);
            _hover = -1;
            Invalidate();
        }

        public void UpdateView(int sel, int top)
        {
            _sel = sel;
            _top = top;
            Invalidate();
        }

        public int VisibleRows => Math.Max(1, Math.Min(_rows, _items.Count - _top));
        public int ListHeight => VisibleRows * _rowH;
        public int TotalHeight => ListHeight + DetailHeight;

        /// <summary>
        /// 窗口该有多大（离屏渲染与自检用；运行期由 PlaceAt 自己算）。
        /// 注意别叫 PreferredSize —— 那是 <see cref="Control"/> 上的属性,同名方法会隐藏它。
        /// </summary>
        public Size DesiredSize() => new Size(ComputeWidth(), TotalHeight + 2);

        /// <summary>当前该不该显示详情行：只有"活动行的名字被省略"时才显示。</summary>
        public string DetailText()
        {
            int active = _hover >= 0 ? _hover : _sel;
            if (active < 0 || active >= _items.Count) return null;
            string name = _items[active] == null ? null : _items[active].Name;
            if (string.IsNullOrEmpty(name)) return null;
            return MeasureText(name) > TextWidth ? name : null;
        }

        public int DetailHeight
        {
            get { return DetailText() == null ? 0 : _font.Height + 4; }
        }

        private int TextWidth => _charW * NameChars;

        private int ComputeWidth()
        {
            int w = PadX + _iconW + IconGap + TextWidth + ScrollW + PadX;
            string detail = DetailText();
            if (detail != null)
            {
                int dw = PadX + MeasureText(detail) + PadX;
                if (dw > w) w = dw;
            }
            return Math.Min(MaxWidth, Math.Max(120, w));
        }

        /// <summary>
        /// 贴着光标显示。(x, top, bottom) 是插入符的屏幕坐标。
        /// 下方放不下就翻到光标上方；若与 VBE 的参数信息窗重叠，整体下移到它下方。
        /// </summary>
        public void PlaceAt(int x, int top, int bottom)
        {
            int w = ComputeWidth();
            int detailH = DetailHeight;
            int listH = VisibleRows * _rowH;
            int h = listH + detailH;

            var scr = Screen.FromPoint(new Point(x, Math.Max(0, bottom)));
            Rectangle wa = scr.WorkingArea;

            int px = x;
            int py = bottom + 2;
            if (py + h > wa.Bottom) py = top - h - 2;

            // ★避让 VBE 的提示窗（参数信息 / 成员列表）：只在"水平也相交"时才躲。
            var rects = Native.VbePopupRects();
            if (rects.Length > 0)
            {
                for (int guard = 0; guard < PopupTryMax; guard++)
                {
                    bool moved = false;
                    foreach (var r in rects)
                    {
                        var me = new Native.RECT { Left = px, Top = py, Right = px + w, Bottom = py + h };
                        if (!Native.Overlap(me, r)) continue;
                        py = r.Bottom + PopupGap;
                        moved = true;
                    }
                    if (!moved) break;
                }
                if (py + h > wa.Bottom)
                {
                    int alt = top - h - PopupGap;
                    bool clash = false;
                    foreach (var r in rects)
                    {
                        var me = new Native.RECT { Left = px, Top = alt, Right = px + w, Bottom = alt + h };
                        if (Native.Overlap(me, r)) { clash = true; break; }
                    }
                    py = (!clash && alt >= wa.Top) ? alt : Math.Max(wa.Top, wa.Bottom - h);
                }
            }

            if (py < wa.Top) py = wa.Top;
            if (px + w > wa.Right) px = wa.Right - w;
            if (px < wa.Left) px = wa.Left;

            // 位置没变就别再 SetWindowPos（每帧都调会让窗口轻微闪）
            if (_lastX == px && _lastY == py && _lastW == w && _lastH == h && Visible) return;
            _lastX = px; _lastY = py; _lastW = w; _lastH = h;

            Native.SetWindowPos(Handle, Native.HWND_TOPMOST, px, py, w, h, Native.SWP_NOACTIVATE);
            Width = w; Height = h;
        }

        public void ResetPlacement()
        {
            _lastX = int.MinValue; _lastY = int.MinValue;
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            // 自绘，避免闪烁
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            int visible = VisibleRows;
            int listH = visible * _rowH;
            string detail = DetailText();

            using (var b = new SolidBrush(CBg)) g.FillRectangle(b, 0, 0, Width, listH);
            if (detail != null)
                using (var b = new SolidBrush(CDetailBg)) g.FillRectangle(b, 0, listH, Width, Height - listH);

            for (int i = 0; i < visible; i++)
            {
                int idx = _top + i;
                if (idx >= _items.Count) break;
                var c = _items[idx];
                if (c == null) continue;
                int y = i * _rowH;
                bool on = idx == _sel;
                bool hovered = idx == _hover;

                if (on) using (var b = new SolidBrush(CSelBg)) g.FillRectangle(b, 0, y, Width, _rowH);
                else if (hovered) using (var b = new SolidBrush(CHoverBg)) g.FillRectangle(b, 0, y, Width, _rowH);

                Color fg = on ? CSelFg : CFg;
                Color hit = on ? CSelHit : CHit;

                int ix = PadX;
                int iy = y + (_rowH - _iconW) / 2;
                DrawIcon(g, c.Kind, ix, iy, _iconW, on);

                int tx = ix + _iconW + IconGap;
                DrawName(g, c, tx, y, fg, hit);
            }

            // 详情行：完整名字（列表里被省略时看这里）
            if (detail != null)
            {
                using (var pen = new Pen(CDetailSep)) g.DrawLine(pen, 0, listH, Width, listH);
                using (var b = new SolidBrush(CDetailFg))
                {
                    var rf = new RectangleF(PadX, listH + (Height - listH - _font.Height) / 2f,
                        Width - PadX * 2, _font.Height + 2);
                    using (var sf = new StringFormat { Trimming = StringTrimming.EllipsisCharacter })
                        g.DrawString(detail, _font, b, rf, sf);
                }
            }

            // 滚动条（只有真需要时才画）
            if (_items.Count > _rows)
            {
                int trackX = Width - ScrollW;
                using (var b = new SolidBrush(CTrack)) g.FillRectangle(b, trackX, 0, ScrollW, listH);
                int total = _items.Count;
                int thumbH = Math.Max(12, (int)((long)listH * Math.Min(_rows, total) / total));
                int maxTop = Math.Max(1, total - _rows);
                int thumbY = (int)((long)(listH - thumbH) * Math.Min(_top, maxTop) / maxTop);
                using (var b = new SolidBrush(CThumb))
                    g.FillRectangle(b, trackX + 1, thumbY, ScrollW - 2, thumbH);
            }

            using (var pen = new Pen(CBorder)) g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        }

        /// <summary>
        /// 按"命中/未命中"把名字切成若干段再上色（与 Python 版 split_by_hits 同一语义）。
        /// 相邻同类字符合并成一段 —— 一个名字通常只有 2~4 段，绘制开销可以忽略。
        /// </summary>
        public static List<KeyValuePair<string, bool>> SplitByHits(string name, int[] positions)
        {
            var outp = new List<KeyValuePair<string, bool>>();
            if (string.IsNullOrEmpty(name)) return outp;
            var hit = new HashSet<int>();
            if (positions != null) foreach (var p in positions) hit.Add(p);

            int i = 0, n = name.Length;
            while (i < n)
            {
                bool isHit = hit.Contains(i);
                int j = i + 1;
                while (j < n && hit.Contains(j) == isHit) j++;
                outp.Add(new KeyValuePair<string, bool>(name.Substring(i, j - i), isHit));
                i = j;
            }
            return outp;
        }

        private void DrawName(Graphics g, Candidate c, int x, int y, Color fg, Color hit)
        {
            string name = c.Name ?? "";
            int baselineY = y + (_rowH - _font.Height) / 2;

            string disp = Truncate(name, TextWidth, out int prefixLen);
            var hits = new List<int>();
            if (c.Hits != null)
                foreach (var p in c.Hits) if (p >= 0 && p < prefixLen) hits.Add(p);

            if (hits.Count == 0)
            {
                TextRenderer.DrawText(g, disp, _font, new Point(x, baselineY), fg,
                    TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
                return;
            }

            int cx = x;
            foreach (var seg in SplitByHits(disp, hits.ToArray()))
                cx = DrawChunk(g, seg.Key, cx, baselineY, seg.Value ? hit : fg);
        }

        private int DrawChunk(Graphics g, string s, int x, int y, Color color)
        {
            if (string.IsNullOrEmpty(s)) return x;
            TextRenderer.DrawText(g, s, _font, new Point(x, y), color,
                TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
            return x + MeasureText(s);
        }

        private int MeasureText(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            return TextRenderer.MeasureText(s, _font, new Size(int.MaxValue, _rowH),
                TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix).Width;
        }

        /// <summary>
        /// 超过 maxW 就用 ... 省略。prefixLen 是"原名字里还能完整显示几个字符"
        /// （命中下标超出这个范围的不再上色 —— 它们已经被省略号吃掉了）。
        /// 先按等宽字体的字符宽估算再微调：逐字符 MeasureText 在长名字上是 O(n)
        /// 次 GDI 调用，而这一步每帧每行都要做。
        /// </summary>
        public string Truncate(string name, int maxW, out int prefixLen)
        {
            prefixLen = name == null ? 0 : name.Length;
            if (string.IsNullOrEmpty(name)) return "";
            if (MeasureText(name) <= maxW) return name;

            int budget = Math.Max(0, maxW - MeasureText(Ellipsis));
            int cut = Math.Min(name.Length, Math.Max(0, budget / Math.Max(1, _charW)));
            while (cut > 0 && MeasureText(name.Substring(0, cut)) > budget) cut--;
            if (cut == 0 && name.Length > 0) cut = 1;

            prefixLen = cut;
            return name.Substring(0, cut) + Ellipsis;
        }

        /// <summary>
        /// 五档图标全部现画（不发位图）：函数=绿箭头、属性=紫方块、变量=蓝竖条、
        /// 类型=琥珀菱形、关键字=灰横线。**拿不到档位就不画**。
        /// </summary>
        private static void DrawIcon(Graphics g, string kind, int x, int y, int w, bool on)
        {
            if (string.IsNullOrEmpty(kind)) return;
            Color color;
            if (!(on ? IconColorOn : IconColor).TryGetValue(kind, out color)) return;

            using (var b = new SolidBrush(color))
            using (var pen = new Pen(color))
            {
                switch (kind)
                {
                    case "func":   // 绿箭头（VBE 的方法图标就是一支箭头）
                        {
                            var pts = new[]
                            {
                                new Point(x, y + w / 4),
                                new Point(x + w / 2, y + w / 4),
                                new Point(x + w / 2, y),
                                new Point(x + w, y + w / 2),
                                new Point(x + w / 2, y + w),
                                new Point(x + w / 2, y + w * 3 / 4),
                                new Point(x, y + w * 3 / 4),
                            };
                            g.FillPolygon(b, pts);
                            break;
                        }
                    case "prop":   // 紫色实心方块
                        g.FillRectangle(b, x + 1, y + w / 4, w - 2, w / 2);
                        break;
                    case "type":   // 琥珀菱形
                        {
                            var pts = new[]
                            {
                                new Point(x + w / 2, y),
                                new Point(x + w, y + w / 2),
                                new Point(x + w / 2, y + w),
                                new Point(x, y + w / 2),
                            };
                            g.FillPolygon(b, pts);
                            break;
                        }
                    case "kw":     // 灰色横线
                        g.FillRectangle(b, x + 1, y + w / 2 - 1, w - 2, 2);
                        break;
                    default:       // var：蓝色小竖条
                        g.FillRectangle(b, x + w / 3, y + 1, Math.Max(2, w / 4), w - 2);
                        break;
                }
            }
        }

        // ------------------------------------------------------------------
        // 鼠标
        // ------------------------------------------------------------------

        private int RowAt(int y)
        {
            if (y < 0 || y >= ListHeight) return -1;
            int i = y / Math.Max(1, _rowH);
            int idx = _top + i;
            return idx < _items.Count ? idx : -1;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_dragScroll)
            {
                int listH = ListHeight;
                int total = _items.Count;
                int maxTop = Math.Max(0, total - _rows);
                if (maxTop > 0 && listH > 0)
                {
                    int thumbH = Math.Max(12, (int)((long)listH * Math.Min(_rows, total) / total));
                    int span = Math.Max(1, listH - thumbH);
                    int y = Math.Max(0, Math.Min(span, e.Y - _dragOffset));
                    int top = (int)((long)y * maxTop / span);
                    ScrollTo?.Invoke(top);
                }
                return;
            }

            // 滚动条区域不参与悬停高亮
            int hover = e.X >= Width - ScrollW && _items.Count > _rows ? -1 : RowAt(e.Y);
            if (hover != _hover)
            {
                _hover = hover;
                Invalidate();
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (_hover != -1) { _hover = -1; Invalidate(); }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            int listH = ListHeight;
            if (e.X >= Width - ScrollW && _items.Count > _rows && e.Y < listH)
            {
                int total = _items.Count;
                int thumbH = Math.Max(12, (int)((long)listH * Math.Min(_rows, total) / total));
                int maxTop = Math.Max(1, total - _rows);
                int thumbY = (int)((long)(listH - thumbH) * Math.Min(_top, maxTop) / maxTop);
                if (e.Y >= thumbY && e.Y <= thumbY + thumbH)
                {
                    _dragScroll = true;
                    _dragOffset = e.Y - thumbY;
                }
                else
                {
                    // 点空白轨道：翻一页
                    Wheel?.Invoke(e.Y < thumbY ? -PageRows : PageRows);
                }
                return;
            }

            int i = RowAt(e.Y);
            if (i >= 0) RowClicked?.Invoke(i);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _dragScroll = false;
        }

        private int PageRows => Math.Max(1, _rows - 1);

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            Wheel?.Invoke(e.Delta > 0 ? -3 : 3);
        }

        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            const int WM_MOUSEACTIVATE = 0x0021;
            const int MA_NOACTIVATE = 3;
            if (m.Msg == WM_MOUSEACTIVATE)
            {
                m.Result = (IntPtr)MA_NOACTIVATE;
                return;
            }
            base.WndProc(ref m);
        }
    }
}
