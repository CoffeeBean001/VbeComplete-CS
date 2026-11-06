// 设置界面 / 词条编辑窗共用的观感：配色、按钮、复选框、分组标题。
// ★与 PopupForm（无边框浮窗）和 SnippetForm（独立窗）同一套色板，它们各自写了一份
//   私有常量；这边新加的两个窗口共用一个，免得再来第三、第四份。
// ★字体**不由这里创建**：Control 不拥有 Font，谁 new 的谁在 Dispose 里放掉，
//   集中创建等于把释放责任藏起来。

using System;
using System.Drawing;
using System.Windows.Forms;

namespace VbeComplete
{
    internal static class UiTheme
    {
        public static readonly Color Body = Hex("#F5F7FA");
        public static readonly Color Card = Hex("#FFFFFF");
        public static readonly Color Line = Hex("#DCE3EC");
        public static readonly Color Text = Hex("#1F2937");
        public static readonly Color Muted = Hex("#6B7280");
        public static readonly Color Accent = Hex("#2563EB");
        public static readonly Color AccentDark = Hex("#1D4ED8");
        public static readonly Color SelectBg = Hex("#E8F0FE");
        public static readonly Color Err = Hex("#B91C1C");
        public static readonly Color Ok = Hex("#15803D");

        public static Color Hex(string s)
        {
            try { return ColorTranslator.FromHtml(s); } catch { return Color.Black; }
        }

        public static Button Primary(string text)
        {
            var b = new Button
            {
                Text = text,
                AutoSize = false,
                FlatStyle = FlatStyle.Flat,
                BackColor = Accent,
                ForeColor = Color.White,
                Size = new Size(96, 32),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false,
            };
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = AccentDark;
            return b;
        }

        public static Button Plain(string text, int w = 96)
        {
            var b = new Button
            {
                Text = text,
                AutoSize = false,
                FlatStyle = FlatStyle.Flat,
                BackColor = Card,
                ForeColor = Text,
                Size = new Size(w, 32),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false,
            };
            b.FlatAppearance.BorderColor = Line;
            b.FlatAppearance.MouseOverBackColor = SelectBg;
            return b;
        }

        public static CheckBox Check(string text, bool on, bool enabled = true)
            => new CheckBox
            {
                Text = text,
                Checked = on,
                AutoSize = true,
                Enabled = enabled,
                ForeColor = Text,
                // ★不写 Transparent：父容器（TableLayoutPanel）在 DrawToBitmap /
                //   离屏渲染时不一定把背景画给子控件，透明标签会露出一块白。
                //   这边整页都是同一个平色，直接填实最稳。
                BackColor = Body,
                FlatStyle = FlatStyle.System,
            };

        /// <summary>
        /// 普通标签。`font` 只在给了的时候才赋值 —— Control.Font 不接受 null。
        /// ★换行标签必须 AutoSize=true + MaximumSize(宽, 0)：只设 MaximumSize 而
        ///   AutoSize=false 时，高度会被"最大高度 0"直接压成 0（整行说明消失，
        ///   截图里只剩一个 100x0 的空标签）。
        /// </summary>
        public static Label Label(string text, Color color, Font font = null, bool wrap = false)
        {
            var l = new Label
            {
                Text = text,
                AutoSize = true,
                ForeColor = color,
                BackColor = Body,
            };
            if (wrap) l.MaximumSize = new Size(560, 0);
            if (font != null) l.Font = font;
            return l;
        }
    }
}
