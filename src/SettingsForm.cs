// VbeComplete 设置界面（托盘右键 →「设置…」/ 双击托盘）
//
// 五页：常用 / 快捷键 / 候选 / 外观 / 词库。
// ★为什么要有这个窗口：以前所有配置都在 settings.ini 里（上百行注释），用户唯一
//   的入口是"用记事本找到那一行、删掉行首的分号"。而两个词库更麻烦 —— 得装
//   DB Browser 才能加一个字，于是"支持自定义"实际上等于没有。
//
// ★保存语义：
//   * 只有**改过**的项才写进 settings.ini（没改的一律不落盘），ini 不会因为点了
//     一次保存就变成一坨；
//   * 「恢复默认」= 把 ini 里这些项重新注释掉（模板本身就是"注释掉 = 用默认值"），
//     不是删行 —— 那一行上面的中文说明要留着；
//   * 绝大多数项是**即时生效**的：Settings 每次都是现读，保存后回调 App.ReloadSettings()
//     让钩子侧那几个缓存字段重读一遍。只有 POPUP_FONT / 配色会重建候选窗样式。
//
// ★词库页直接读写那两个 SQLite 库（经 WordBank），改完立刻 Reload + 清候选池，
//   不重启也能用。库路径可能是用户自己配的（KEYWORD_DB / FUNCTION_DB），
//   所以界面上永远显示**真实路径**，别让用户以为在改项目目录里那一份。

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace VbeComplete
{
    internal sealed class SettingsForm : Form
    {
        // ---- 一项配置在界面上的样子 ----
        private sealed class Row
        {
            public string Key;
            public string Def;                 // 默认值（字符串形式）
            public Func<string> Read;          // 从控件读出当前值
            public Action<string> Write;       // 把值写回控件（"恢复默认"用）
            public Func<bool> Checked;         // 勾选框此刻勾着没有（非勾选项为 null）
            public string Original;            // 打开窗口时 ini 里的值（null = 没写过）
        }

        private sealed class BankView
        {
            public string Bank;
            public TabPage Page;
            public TextBox Search;
            public ListView List;
            public List<WordEntry> All = new List<WordEntry>();
            public Label Count;                // "共 N 条，启用 M 条"
            public bool Building;              // 程序性重建列表，别把 Checked 事件当成用户操作
        }

        private static readonly string[] PageNames = { "常用", "快捷键", "候选", "外观", "词库" };

        private readonly Action _onSaved;              // 设置保存后（App 重读 + 重画候选窗）
        private readonly Action _onWordBankChanged;    // 词库改动后（清候选池）

        private readonly Font _uiFont;
        private readonly Font _boldFont;
        private readonly Font _descFont;

        private ListBox _nav;
        private Panel _host;
        private Panel[] _pages;
        private Label _status;

        private readonly List<Row> _rows = new List<Row>();
        private readonly List<BankView> _banks = new List<BankView>();

        public SettingsForm(Action onSaved, Action onWordBankChanged, int startPage = 0)
        {
            _onSaved = onSaved;
            _onWordBankChanged = onWordBankChanged;

            Text = "VbeComplete 设置";
            try { Icon = IconAssets.Window(); } catch { }
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = UiTheme.Body;
            _uiFont = new Font("Microsoft YaHei UI", 9f);
            _boldFont = new Font("Microsoft YaHei UI", 10.5f, FontStyle.Bold);
            _descFont = new Font("Microsoft YaHei UI", 8.25f);
            Font = _uiFont;
            AutoScaleMode = AutoScaleMode.Font;
            MinimumSize = new Size(720, 480);
            Size = new Size(880, 620);
            ShowInTaskbar = true;
            KeyPreview = true;

            BuildUi();
            SelectPage(startPage);

            KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape) Close();
                else if (e.Control && e.KeyCode == Keys.S) SaveAll();
            };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _uiFont?.Dispose();
                _boldFont?.Dispose();
                _descFont?.Dispose();
            }
            base.Dispose(disposing);
        }

        // ==================================================================
        // 骨架
        // ==================================================================

        private void BuildUi()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                BackColor = UiTheme.Body,
                Margin = new Padding(0),
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 156f));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58f));

            // ---- 左侧导航 ----
            _nav = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 44,
                IntegralHeight = false,
                BackColor = UiTheme.Body,
                ForeColor = UiTheme.Text,
                Margin = new Padding(0, 8, 0, 0),
            };
            foreach (var n in PageNames) _nav.Items.Add(n);
            _nav.DrawItem += NavDraw;
            _nav.SelectedIndexChanged += (s, e) => SelectPage(_nav.SelectedIndex);
            root.Controls.Add(_nav, 0, 0);

            // ---- 右侧页面 ----
            _host = new Panel { Dock = DockStyle.Fill, BackColor = UiTheme.Body, Margin = new Padding(0) };
            root.Controls.Add(_host, 1, 0);

            _pages = new Panel[PageNames.Length];
            _pages[0] = BuildCommonPage();
            _pages[1] = BuildHotkeyPage();
            _pages[2] = BuildCandidatePage();
            _pages[3] = BuildLookPage();
            _pages[4] = BuildWordBankPage();
            foreach (var p in _pages) { p.Visible = false; _host.Controls.Add(p); }

            // ---- 底栏 ----
            var bar = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.White,
                Margin = new Padding(0),
                Padding = new Padding(16, 0, 16, 0),
            };
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bar.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            _status = UiTheme.Label("", UiTheme.Muted, _uiFont);
            _status.BackColor = Color.White;         // 底栏是白底
            _status.Anchor = AnchorStyles.Left;
            bar.Controls.Add(_status, 0, 0);

            var btns = new FlowLayoutPanel
            {
                AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.RightToLeft,
                BackColor = UiTheme.Body, Anchor = AnchorStyles.Right, Margin = new Padding(0),
            };
            var save = UiTheme.Primary("保存");
            save.Size = new Size(110, 34);
            save.Click += (s, e) => SaveAll();
            var reset = UiTheme.Plain("恢复默认", 110);
            reset.Size = new Size(110, 34);
            reset.Click += (s, e) => ResetDefaults();
            btns.Controls.Add(save);
            btns.Controls.Add(reset);
            bar.Controls.Add(btns, 1, 0);

            // 底栏横跨两列
            root.SetColumnSpan(bar, 2);
            root.Controls.Add(bar, 0, 1);

            Controls.Add(root);
        }

        private void NavDraw(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            bool sel = (e.State & DrawItemState.Selected) != 0;
            using (var b = new SolidBrush(sel ? UiTheme.SelectBg : UiTheme.Body))
                e.Graphics.FillRectangle(b, e.Bounds);
            if (sel)
                using (var b = new SolidBrush(UiTheme.Accent))
                    e.Graphics.FillRectangle(b, new Rectangle(e.Bounds.X, e.Bounds.Y + 9, 3,
                        Math.Max(6, e.Bounds.Height - 18)));
            var font = sel ? _boldFont : _uiFont;
            TextRenderer.DrawText(e.Graphics, Convert.ToString(_nav.Items[e.Index]), font,
                new Rectangle(e.Bounds.X + 18, e.Bounds.Y, e.Bounds.Width - 18, e.Bounds.Height),
                sel ? UiTheme.Accent : UiTheme.Text,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
        }

        /// <summary>从外面切页（托盘「添加关键字」直接落到词库页）。</summary>
        public void GoPage(int i) => SelectPage(i);

        private void SelectPage(int i)
        {
            if (i < 0 || i >= _pages.Length || _pages[i] == null) return;
            for (int k = 0; k < _pages.Length; k++)
                if (_pages[k] != null) _pages[k].Visible = k == i;
            if (_nav.SelectedIndex != i) _nav.SelectedIndex = i;
            if (i == 4) ReloadBank(_banks[0]);       // 进词库页时刷一次（外面可能改过库）
        }

        // ==================================================================
        // 行构件
        // ==================================================================

        private static TableLayoutPanel Page()
        {
            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                BackColor = UiTheme.Body,
                AutoScroll = true,
                Padding = new Padding(20, 10, 20, 14),
            };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            return t;
        }

        private static void Add(TableLayoutPanel page, Control c)
        {
            int i = page.RowCount;
            page.RowCount = i + 1;
            page.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            page.Controls.Add(c, 0, i);
        }

        /// <summary>小节标题（一行粗体 + 上面留白）。</summary>
        private Label Section(string text)
        {
            var l = UiTheme.Label(text, UiTheme.Text, _boldFont);
            l.Margin = new Padding(0, 16, 0, 8);
            return l;
        }

        /// <summary>勾选项：标题一行、说明一行（说明自动换行，不会被切掉）。</summary>
        /// <param name="invert">
        /// 反向项：ini 键名是 `NO_XXX`（**写了 = 把这个功能关掉**），而界面上那句话说的是
        /// 功能本身，所以勾上必须对应"ini 里没有这一项"。
        /// ★v176 的教训：之前一律 `勾上 = 写 1`，于是"自动配对括号 / 引号"这类项的框
        ///   默认是空的（功能明明开着）、勾上反倒是关掉 —— 界面上每一个 NO_ 项都和用户
        ///   的心智反着来。凡"键名带 NO_ / 标签说的是功能而不是禁用"的，一律走这个口子。
        /// </param>
        private void AddToggle(TableLayoutPanel page, string key, bool def,
            string title, string desc, bool invert = false)
        {
            var box = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 2,                     // ★必须显式写：RowCount 默认 0 时
                AutoSize = true,                  //   加到 (0,1) 的那一行不会占高度，
                AutoSizeMode = AutoSizeMode.GrowAndShrink,   //   说明文字就整行看不见
                BackColor = UiTheme.Body,
                Margin = new Padding(0, 0, 0, 10),
            };
            box.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            box.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            box.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var cb = UiTheme.Check(title, invert ? !Settings.Flag(key, def) : Settings.Flag(key, def));
            cb.Margin = new Padding(0, 0, 0, 1);
            box.Controls.Add(cb, 0, 0);

            var lb = UiTheme.Label(desc, UiTheme.Muted, _descFont, true);
            lb.Margin = new Padding(20, 0, 0, 0);
            box.Controls.Add(lb, 0, 1);

            Add(page, box);
            _rows.Add(new Row
            {
                Key = key,
                Def = def ? "1" : "0",
                Original = Settings.IniValue(key),
                // 反向项：勾上 = 不写这个 NO_ 键（"0" 与"没写过"等价，SaveAll 会跳过默认值，
                // 所以来回点也不会把 ini 塞满）。
                Read = () => (cb.Checked ^ invert) ? "1" : "0",
                Write = v =>
                {
                    bool on = v != "0";
                    cb.Checked = invert ? !on : on;
                },
                Checked = () => cb.Checked,
            });
        }

        /// <summary>反向项（键名 `NO_XXX`，勾上 = 功能开着）。见 <see cref="AddToggle"/>。</summary>
        private void AddToggleOff(TableLayoutPanel page, string key, bool def,
            string title, string desc)
            => AddToggle(page, key, def, title, desc, invert: true);

        /// <summary>带输入框的一行：左边标题 + 说明，右边控件。</summary>
        private void AddField(TableLayoutPanel page, string key, string def,
            string title, string desc, Control value, Func<string> read, Action<string> write)
        {
            var row = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = UiTheme.Body,
                Margin = new Padding(0, 0, 0, 10),
            };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var left = new TableLayoutPanel
            {
                ColumnCount = 1, RowCount = 2,     // 同上：不写 RowCount 说明行会消失
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = UiTheme.Body, Margin = new Padding(0),
            };
            left.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var t = UiTheme.Label(title, UiTheme.Text, _uiFont);
            t.Margin = new Padding(0, 0, 0, 1);
            left.Controls.Add(t, 0, 0);
            if (!string.IsNullOrEmpty(desc))
            {
                var d = UiTheme.Label(desc, UiTheme.Muted, _descFont, true);
                d.Margin = new Padding(0);
                left.Controls.Add(d, 0, 1);
            }
            row.Controls.Add(left, 0, 0);

            value.Anchor = AnchorStyles.Right;
            value.Margin = new Padding(12, 2, 0, 0);
            row.Controls.Add(value, 1, 0);

            Add(page, row);
            _rows.Add(new Row
            {
                Key = key,
                Def = def,
                Original = Settings.IniValue(key),
                Read = read,
                Write = write,
            });
        }

        private NumericUpDown Number(int min, int max, int value, int width = 90)
            => new NumericUpDown
            {
                Minimum = min, Maximum = max, Value = Math.Max(min, Math.Min(max, value)),
                Width = width, TextAlign = HorizontalAlignment.Right,
                BackColor = UiTheme.Card, ForeColor = UiTheme.Text,
                BorderStyle = BorderStyle.FixedSingle,
            };

        private ComboBox Combo(string[] items, int index, int width)
        {
            var c = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList, Width = width,
                BackColor = UiTheme.Card, ForeColor = UiTheme.Text,
            };
            c.Items.AddRange(items);
            c.SelectedIndex = Math.Max(0, Math.Min(items.Length - 1, index));
            return c;
        }

        // ==================================================================
        // 第 1 页：常用
        // ==================================================================

        private Panel BuildCommonPage()
        {
            var page = Page();

            Add(page, Section("补全行为"));
            AddToggle(page, "AUTO_SPACE", true, "关键字后自动补空格",
                "敲完 `End If` 补一个空格。关掉之后补全只替换词本身。");
            AddToggle(page, "AUTO_PAREN", true, "函数名后自动补一对括号",
                "补全 `MsgBox` 得到 `MsgBox()`；类型名（String / Long）永远不补括号。");
            AddToggle(page, "REQUIRE_KEY_TRIGGER", true, "只有真的在打字时才自动弹候选",
                "鼠标右键粘贴、拖拽、别的程序改代码造成的文本变化只收窗、不弹窗。");
            AddToggle(page, "HIDE_ON_COMPLETE", true, "词打全了就收窗",
                "例如已经打完 `Select` 时不再挂在屏幕上。Ctrl+Space 不受影响。");

            Add(page, Section("括号与缩进"));
            AddToggleOff(page, "NO_AUTOPAIR", false, "自动配对括号 / 引号",
                "敲 `(` 补 `)`，敲 `\"` 补另一个 `\"`。不勾 = 只敲你按的那一个字符。");
            AddToggleOff(page, "NO_IME_PUNCT", false, "中文标点也一起配对",
                "输入法切在中文标点 / 全角档时，自动配出来的另一半也跟着全角" +
                "（敲 `(` 得到 `（）`，引号得到 `“”`），和你此刻打出来的字一致。" +
                "不勾 = 不管输入法什么档，一律补半角 `()` / `\"\"`。");
            AddToggleOff(page, "NO_ENTER_INDENT", false, "回车自动缩进",
                "回车后新行沿用上一行的缩进，`If` / `For` / `Sub` 这类块头自动再进一级。" +
                "不勾 = 回车完全交还给 VBE 原生行为。");
            AddToggleOff(page, "NO_AUTOCLOSE", false, "回车自动补结构收尾",
                "`Sub` / `If` / `For` 回车后自动补 `End Sub` / `End If` / `Next`，光标落在两行中间。" +
                "不勾 = 只缩进、不补收尾。");
            AddToggleOff(page, "NO_BRANCH_ALIGN", false, "分支缩进对齐",
                "`Else` / `Case` 自动对齐到所属块的层级。不勾 = 按你敲的列原样放着。");

            var unit = Combo(new[] { "自动（看工程里既有的写法）", "2 个空格", "4 个空格", "Tab" },
                IndentIndex(Settings.Raw("INDENT_UNIT")), 220);
            AddField(page, "INDENT_UNIT", "", "一级缩进",
                "留空 / 自动 = 按工程里已有的缩进风格判断。",
                unit,
                () => new[] { "", "2", "4", "tab" }[Math.Max(0, unit.SelectedIndex)],
                v =>
                {
                    int i = Array.IndexOf(new[] { "", "2", "4", "tab" }, v ?? "");
                    unit.SelectedIndex = i < 0 ? 0 : i;
                });

            AddToggleOff(page, "NO_BS_DELLINE", false, "退格一次删掉整行空白",
                "行首只有缩进时，按一下退格就删掉整段空白。不勾 = 一格一格退。");
            AddToggleOff(page, "NO_BS_DELPAIR", false, "退格一次删掉成对的空括号 / 引号",
                "`()` 中间按退格，两个一起删。不勾 = 只删前一个字符，留下一个孤零零的 `）`。");

            Add(page, Section("排障"));
            AddToggle(page, "LOG", false, "写诊断日志",
                "正常使用建议关着（关着时零开销）。排查问题再打开，日志在项目目录的 VbeComplete.log。");
            AddToggle(page, "NO_WORDDB", false, "跳过两个词库，只用内置种子",
                "只在怀疑词库被改坏时临时打开。打开后你加的词条一条都不会出现。");

            return Wrap(page);
        }

        private static int IndentIndex(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return 0;
            switch (raw.Trim().ToLowerInvariant())
            {
                case "2": return 1;
                case "4": return 2;
                case "tab": return 3;
                default: return 0;
            }
        }

        // ==================================================================
        // 第 2 页：快捷键
        // ==================================================================

        private Panel BuildHotkeyPage()
        {
            var page = Page();

            Add(page, Section("自己的组合键"));
            AddToggleOff(page, "NO_SNIPPETS", false, "Ctrl+Q 打开代码片段窗",
                "在 VBE 代码窗里按 Ctrl+Q。不勾 = 把这个键还给 VBE。");
            AddToggleOff(page, "NO_COMMENT", false, "Ctrl+/ 批量注释 / 取消注释",
                "走 VBE 自己的命令，Undo 归 VBE。不勾 = 还给 VBE。");
            AddToggleOff(page, "NO_CTRL_NAV", false, "Ctrl+↑ / Ctrl+↓ 跳过程两端",
                "Ctrl+↑ 跳到本过程最后一个 Dim 行，Ctrl+↓ 跳到收尾行上方最后一条代码。");
            AddToggleOff(page, "NO_SHIFT_NAV", false, "Shift+↑/↓ 跳出候选列表并上下移一行",
                "候选窗开着时，先按 Shift+方向键关掉候选、再按方向键正常移光标。" +
                "不勾 = 交还 VBE（效果是选中文本）。");

            Add(page, Section("行操作"));
            AddToggleOff(page, "NO_DUP_LINE", false, "Ctrl+D 复制当前行",
                "不勾 = 还给 VBE。");
            AddToggleOff(page, "NO_DEL_LINE", false, "Ctrl+Shift+K 删除当前行",
                "不勾 = 还给 VBE。");
            AddToggleOff(page, "NO_INSERT_ABOVE", false, "Ctrl+Enter 在上方插入一行",
                "不勾 = 还给 VBE（回车就是普通回车）。");

            return Wrap(page);
        }

        // ==================================================================
        // 第 3 页：候选
        // ==================================================================

        private Panel BuildCandidatePage()
        {
            var page = Page();

            Add(page, Section("候选里放什么"));
            AddToggle(page, "VBA_BUILTINS", true, "VBA 内建函数与内建类型",
                "`MsgBox` / `Split` / `String` / `Long` 这些。");
            AddToggle(page, "VBA_KEYWORDS", true, "VBA 语言关键字",
                "`If` / `For` / `Dim` 这些。关掉 = 关键字一条都不提示，函数与类型照旧。");
            AddToggle(page, "VBA_CONSTANTS", false, "VBA 内建枚举常量",
                "`vbCrLf` / `vbYes` 这些。默认关：数量大、噪音多。");
            AddToggle(page, "HOST_ENUMS", false, "宿主类型库枚举常量",
                "`xlUp` / `msoTrue` 这些。默认关，理由同上。");

            Add(page, Section("什么位置弹"));
            AddToggle(page, "MEMBER_POPUP", false, "`对象.` 之后也弹候选",
                "不勾 = 一个都不出，整个让给 VBE 自己的成员列表（它的类型信息更准）。" +
                "勾上 = 出：能定出类型的按类型收窄。Ctrl+Space 不受这个开关影响。");
            AddToggle(page, "DECL_POPUP", false, "声明位置（正在起新名字）也弹候选",
                "`Dim x As |` 这种位置默认不弹。");

            var rows = Number(3, 30, Settings.Int("VIEW_ROWS", 15, 3, 30));
            AddField(page, "VIEW_ROWS", "15", "候选窗最多显示几行",
                "超过就滚动。屏幕小可以调到 8~10。", rows,
                () => ((int)rows.Value).ToString(),
                v => { int n; if (int.TryParse(v, out n)) rows.Value = Math.Max(3, Math.Min(30, n)); });

            return Wrap(page);
        }

        // ==================================================================
        // 第 4 页：外观
        // ==================================================================

        private Panel BuildLookPage()
        {
            var page = Page();

            Add(page, Section("字体"));
            var faces = new List<string> { "自动（跟 VBE 编辑器，小一号）" };
            try
            {
                // 只列等宽字体：候选列表里参差不齐的宽度比 VBE 还乱。
                // ★FontFamily 没有"我是不是等宽"的查询接口，只能按名字认（下面这几个词
                //   覆盖了 Windows 上能见到的绝大多数等宽字体）。认不出来也没关系 ——
                //   下拉框是可编辑的，手打字体名照样生效。
                var monoHints = new[] { "mono", "console", "consolas", "courier", "code",
                                        "hack", "menlo", "inconsolata", "iosevka", "fira",
                                        "source code", "noto sans mono", "sarasa", "等距" };
                foreach (var f in FontFamily.Families)
                {
                    try
                    {
                        if (!f.IsStyleAvailable(FontStyle.Regular)) continue;
                        string n = f.Name.ToLowerInvariant();
                        if (monoHints.Any(h => n.Contains(h))) faces.Add(f.Name);
                    }
                    catch { }
                }
            }
            catch { }
            string curFace = Settings.Raw("POPUP_FONT") ?? "";
            int faceIdx = 0;
            for (int i = 1; i < faces.Count; i++)
                if (string.Equals(faces[i], curFace, StringComparison.OrdinalIgnoreCase)) { faceIdx = i; break; }

            var face = Combo(faces.ToArray(), faceIdx, 240);
            face.DropDownStyle = ComboBoxStyle.DropDown;        // 允许手打没列出来的字体
            AddField(page, "POPUP_FONT", "", "候选窗字体",
                "留空 / 自动 = 先跟 VBE 编辑器字体（去掉 Light/Thin），再退到等宽字体。",
                face,
                () => face.SelectedIndex <= 0 ? "" : face.Text.Trim(),
                v =>
                {
                    face.SelectedIndex = string.IsNullOrEmpty(v) ? 0 : -1;
                    face.Text = string.IsNullOrEmpty(v) ? faces[0] : v;
                });

            var size = Number(-6, 24, Settings.Int("POPUP_FONT_SIZE", 0, -6, 24));
            AddField(page, "POPUP_FONT_SIZE", "0", "字号",
                "0 = 比 VBE 小一号（默认）；-2 = 小两号；正数 = 直接用这个磅值。",
                size,
                () => ((int)size.Value).ToString(),
                v => { int n; if (int.TryParse(v, out n)) size.Value = Math.Max(-6, Math.Min(24, n)); });

            Add(page, Section("配色"));
            AddColor(page, "POPUP_SEL_BG", "#0A62B0", "选中行底色",
                "默认是深蓝（和 VBE 里选中文本一个观感）。");
            AddColor(page, "POPUP_SEL_FG", "#FFFFFF", "选中行文字色", "");
            AddColor(page, "POPUP_SEL_HIT", "#FFD97A", "选中行里命中字母的颜色",
                "深蓝底上用亮黄，比黑色加粗更清楚。");

            return Wrap(page);
        }

        private void AddColor(TableLayoutPanel page, string key, string def, string title, string desc)
        {
            string cur = Settings.Raw(key);
            var color = ParseColor(string.IsNullOrWhiteSpace(cur) ? def : cur);
            var btn = new Button
            {
                Text = ToHex(color),
                Size = new Size(110, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = color,
                ForeColor = Contrast(color),
                Margin = new Padding(12, 2, 0, 0),
            };
            btn.FlatAppearance.BorderColor = UiTheme.Line;
            btn.Click += (s, e) =>
            {
                using (var dlg = new ColorDialog { Color = btn.BackColor, FullOpen = true })
                {
                    if (dlg.ShowDialog(this) != DialogResult.OK) return;
                    btn.BackColor = dlg.Color;
                    btn.ForeColor = Contrast(dlg.Color);
                    btn.Text = ToHex(dlg.Color);
                }
            };
            AddField(page, key, def, title, desc, btn,
                () => ToHex(btn.BackColor),
                v =>
                {
                    var c = ParseColor(string.IsNullOrWhiteSpace(v) ? def : v);
                    btn.BackColor = c;
                    btn.ForeColor = Contrast(c);
                    btn.Text = ToHex(c);
                });
        }

        private static Color ParseColor(string s)
        {
            try { return ColorTranslator.FromHtml(s); } catch { return Color.Black; }
        }

        private static string ToHex(Color c) => "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");

        private static Color Contrast(Color c)
            => (c.R * 299 + c.G * 587 + c.B * 114) / 1000 > 140 ? Color.Black : Color.White;

        // ==================================================================
        // 第 5 页：词库
        // ==================================================================

        private Panel BuildWordBankPage()
        {
            var page = Page();
            page.AutoScroll = false;
            page.RowCount = 3;
            page.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 说明
            page.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));  // 两个库
            page.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 路径

            var intro = UiTheme.Label(
                "这两个库就是候选的全部来源。改完立刻生效（不用重启）：新加的词下一次按键就能补出来。" +
                "勾掉最左边的框 = 先停用，但不删。",
                UiTheme.Muted, _descFont, true);
            intro.Margin = new Padding(0, 0, 0, 10);
            page.Controls.Add(intro, 0, 0);

            var tabs = new TabControl { Dock = DockStyle.Fill, Font = _uiFont };
            tabs.TabPages.Add(BuildBankTab(WordBank.Keywords, "关键字", "keywords-cs.db"));
            tabs.TabPages.Add(BuildBankTab(WordBank.Functions, "函数与类型", "functions-cs.db"));
            page.Controls.Add(tabs, 0, 1);

            var paths = UiTheme.Label("", UiTheme.Muted, _descFont, true);
            paths.Name = "paths";
            paths.Margin = new Padding(0, 8, 0, 0);
            paths.Text = "关键字库：" + WordLibrary.KeywordPath() +
                         "\r\n函数库：" + WordLibrary.FunctionPath();
            page.Controls.Add(paths, 0, 2);

            return Wrap(page);
        }

        private TabPage BuildBankTab(string bank, string title, string fileName)
        {
            var tab = new TabPage(title) { BackColor = UiTheme.Body, Padding = new Padding(8) };

            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4,
                BackColor = UiTheme.Body, Margin = new Padding(0),
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 40f));   // 搜索 + 刷新
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));   // 列表
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 44f));   // 按钮
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // 状态

            var top = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1,
                BackColor = UiTheme.Body, Margin = new Padding(0),
            };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var search = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = UiTheme.Card,
                ForeColor = UiTheme.Text,
                PlaceholderText = "搜索 " + fileName + "（含未启用的条目）",
                Margin = new Padding(0, 4, 8, 4),
            };
            top.Controls.Add(search, 0, 0);
            var refresh = UiTheme.Plain("刷新", 76);
            refresh.Margin = new Padding(0, 4, 0, 4);
            top.Controls.Add(refresh, 1, 0);
            grid.Controls.Add(top, 0, 0);

            var list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                CheckBoxes = true,
                FullRowSelect = true,
                HideSelection = false,
                MultiSelect = true,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = UiTheme.Card,
                ForeColor = UiTheme.Text,
                Font = _uiFont,
                Margin = new Padding(0),
            };
            list.Columns.Add("词", 260);
            if (bank == WordBank.Keywords) list.Columns.Add("补完跟空格", 110);
            else
            {
                list.Columns.Add("种类", 100);
                list.Columns.Add("补括号", 90);
            }
            list.Columns.Add("状态", 90);
            grid.Controls.Add(list, 0, 1);

            var bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, WrapContents = false, BackColor = UiTheme.Body,
                Margin = new Padding(0, 8, 0, 0),
            };
            var add = UiTheme.Primary("新增…");
            add.Size = new Size(100, 32);
            var edit = UiTheme.Plain("编辑…", 90);
            var del = UiTheme.Plain("删除", 90);
            var open = UiTheme.Plain("打开库所在文件夹", 150);
            bar.Controls.Add(add);
            bar.Controls.Add(edit);
            bar.Controls.Add(del);
            bar.Controls.Add(open);
            grid.Controls.Add(bar, 0, 2);

            var status = UiTheme.Label("", UiTheme.Muted, _descFont, true);
            status.Margin = new Padding(0, 4, 0, 0);
            grid.Controls.Add(status, 0, 3);

            tab.Controls.Add(grid);

            var view = new BankView
            {
                Bank = bank, Page = tab, Search = search, List = list, Count = status,
            };
            _banks.Add(view);

            // ---- 事件 ----
            search.TextChanged += (s, e) => Refilter(view);
            refresh.Click += (s, e) => ReloadBank(view);
            add.Click += (s, e) => AddWord(view);
            edit.Click += (s, e) => EditWord(view);
            del.Click += (s, e) => DeleteWords(view);
            open.Click += (s, e) => RevealFile(WordBank.PathOf(bank));
            list.DoubleClick += (s, e) => EditWord(view);
            list.ItemChecked += (s, e) =>
            {
                if (view.Building) return;
                // 只写"和内存里那一份不一样"的：否则一次刷新会把整库重写一遍
                var want = new List<string>();
                foreach (ListViewItem it in list.Items)
                {
                    var w = ItemWord(it);
                    var mem = view.All.FirstOrDefault(x => string.Equals(x.Word, w,
                        StringComparison.OrdinalIgnoreCase));
                    if (mem == null) continue;
                    bool now = it.Checked;
                    if (mem.Enabled != now) want.Add(w + (now ? "\u0001" : "\u0000"));
                }
                if (want.Count == 0) return;
                var on = want.Where(x => x.EndsWith("\u0001")).Select(x => x.Substring(0, x.Length - 1)).ToList();
                var off = want.Where(x => x.EndsWith("\u0000")).Select(x => x.Substring(0, x.Length - 1)).ToList();
                string err = "";
                bool ok = true;
                if (on.Count > 0) ok &= WordBank.SetEnabled(bank, on, true, out err);
                if (off.Count > 0) ok &= WordBank.SetEnabled(bank, off, false, out err);
                if (!ok) { SetStatus("改启用状态失败：" + err, true); return; }
                foreach (var x in on) { var m = view.All.FirstOrDefault(y => y.Word == x); if (m != null) m.Enabled = true; }
                foreach (var x in off) { var m = view.All.FirstOrDefault(y => y.Word == x); if (m != null) m.Enabled = false; }
                AfterBankChange("已" + (on.Count > 0 ? "启用 " + on.Count + " 条" : "") +
                                (off.Count > 0 ? "停用 " + off.Count + " 条" : ""));
            };

            return tab;
        }

        private static string ItemWord(ListViewItem it)
            => it.Tag as string ?? (it.Text ?? "").Trim();

        private void ReloadBank(BankView v)
        {
            string err;
            var all = WordBank.All(v.Bank, out err);
            v.All = all;
            Refilter(v);
            if (!string.IsNullOrEmpty(err)) SetStatus("读词库失败：" + err, true);
        }

        private void Refilter(BankView v)
        {
            string q = (v.Search.Text ?? "").Trim();
            var words = string.IsNullOrEmpty(q)
                ? v.All
                : v.All.Where(x => x.Word.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0).ToList();

            v.Building = true;
            v.List.BeginUpdate();
            try
            {
                v.List.Items.Clear();
                foreach (var e in words)
                {
                    var it = new ListViewItem(e.Word) { Tag = e.Word, Checked = e.Enabled };
                    if (v.Bank == WordBank.Keywords)
                        it.SubItems.Add(e.WantsSpace ? "是" : "否");
                    else
                    {
                        it.SubItems.Add(e.IsType ? "类型" : "函数");
                        it.SubItems.Add(e.IsType ? "—" : (e.WantsParens ? "是" : "否"));
                    }
                    it.SubItems.Add(e.Enabled ? "启用" : "已停用");
                    if (!e.Enabled) it.ForeColor = UiTheme.Muted;
                    v.List.Items.Add(it);
                }
            }
            finally
            {
                v.List.EndUpdate();
                v.Building = false;
            }

            int on = v.All.Count(x => x.Enabled);
            string count = string.IsNullOrEmpty(q)
                ? "共 " + v.All.Count + " 条，启用 " + on + " 条"
                : "命中 " + words.Count + " 条（全部 " + v.All.Count + " 条，启用 " + on + " 条）";
            if (v.List.Items.Count > 0) v.List.EnsureVisible(0);
            if (v.Count != null) v.Count.Text = count;
        }

        private void AddWord(BankView v)
        {
            var list = WordEditForm.Ask(this, v.Bank, null);
            if (list == null || list.Count == 0) return;

            var existing = new HashSet<string>(v.All.Select(x => x.Word), StringComparer.OrdinalIgnoreCase);
            int added = 0, updated = 0;
            string err = "";
            foreach (var e in list)
            {
                bool isNew = !existing.Contains(e.Word);
                if (!WordBank.Save(v.Bank, e, null, out err))
                {
                    SetStatus("写入失败：" + err, true);
                    ReloadBank(v);
                    return;
                }
                if (isNew) added++; else updated++;
            }
            ReloadBank(v);
            SearchFor(v, list.Count == 1 ? list[0].Word : null);
            AfterBankChange("新增 " + added + " 条" + (updated > 0 ? "、覆盖 " + updated + " 条" : ""));
        }

        private void EditWord(BankView v)
        {
            if (v.List.SelectedItems.Count == 0) { SetStatus("先在列表里选一条。", true); return; }
            var word = ItemWord(v.List.SelectedItems[0]);
            var mem = v.All.FirstOrDefault(x => string.Equals(x.Word, word,
                StringComparison.OrdinalIgnoreCase));
            if (mem == null) { SetStatus("这一条刚从库里被别人删掉了，刷新一下。", true); return; }

            var list = WordEditForm.Ask(this, v.Bank, mem);
            if (list == null || list.Count == 0) return;
            string err;
            if (!WordBank.Save(v.Bank, list[0], mem.Word, out err))
            {
                SetStatus("写入失败：" + err, true);
                return;
            }
            ReloadBank(v);
            SearchFor(v, list[0].Word);
            AfterBankChange("已保存「" + list[0].Word + "」");
        }

        private void DeleteWords(BankView v)
        {
            if (v.List.SelectedItems.Count == 0) { SetStatus("先在列表里选一条（可多选）。", true); return; }
            var words = v.List.SelectedItems.Cast<ListViewItem>().Select(ItemWord).ToList();
            var ask = words.Count == 1
                ? "从词库里删掉「" + words[0] + "」？\r\n\r\n删掉之后它不会再出现在候选里。"
                : "从词库里删掉这 " + words.Count + " 条？";
            if (InfoBox.Confirm(ask, "删除词条") != DialogResult.OK) return;

            string err;
            int n = WordBank.Delete(v.Bank, words, out err);
            if (!string.IsNullOrEmpty(err)) { SetStatus("删除失败：" + err, true); return; }
            ReloadBank(v);
            AfterBankChange("删掉 " + n + " 条");
        }

        private void SearchFor(BankView v, string word)
        {
            if (string.IsNullOrEmpty(word)) return;
            foreach (ListViewItem it in v.List.Items)
                if (string.Equals(ItemWord(it), word, StringComparison.OrdinalIgnoreCase))
                {
                    it.Selected = true;
                    it.Focused = true;
                    it.EnsureVisible();
                    return;
                }
        }

        /// <summary>写完库的统一收尾：重灌内存镜像 + 清候选池 + 报状态。</summary>
        private void AfterBankChange(string what)
        {
            WordBank.Reload();
            try { _onWordBankChanged?.Invoke(); } catch (Exception e) { Log.Write("词库变更回调失败：{0}", e.Message); }
            SetStatus(what + " —— 已经生效，不用重启。", false);
        }

        private static void RevealFile(string path)
        {
            try
            {
                if (!File.Exists(path))
                {
                    InfoBox.Show("词库文件不存在", "路径：" + path);
                    return;
                }
                Process.Start(new ProcessStartInfo("explorer.exe", "/select,\"" + path + "\"")
                { UseShellExecute = true });
            }
            catch (Exception e) { Log.Write("打开词库所在文件夹失败：{0}", e.Message); }
        }

        // ==================================================================
        // 保存 / 恢复默认
        // ==================================================================

        private void SaveAll()
        {
            int n = 0;
            var failed = new List<string>();
            foreach (var r in _rows)
            {
                string now;
                try { now = r.Read(); } catch { continue; }
                bool atDefault = string.Equals(now, r.Def, StringComparison.Ordinal);
                if (string.Equals(now, r.Original ?? r.Def, StringComparison.Ordinal) &&
                    r.Original != null)
                    continue;                                  // 没动过
                if (r.Original == null && atDefault)
                    continue;                                  // 仍然是默认值，别往 ini 里塞
                // 改回默认值 = 把那一行重新注释掉（模板里的中文说明和默认值都留着），
                // 不然来回点几次就会在 ini 里攒下一堆 `; NO_XXX = 0` 这样的活行。
                string target = atDefault ? null : now;
                bool ok = target == null ? Settings.Reset(r.Key, r.Def) : Settings.Set(r.Key, target);
                if (ok) { r.Original = target; n++; }
                else failed.Add(r.Key);
            }

            if (failed.Count > 0)
            {
                SetStatus("写 settings.ini 失败：" + string.Join(", ", failed) +
                          "（" + Settings.LastError() + "）", true);
                return;
            }

            try { _onSaved?.Invoke(); } catch (Exception e) { Log.Write("设置保存回调失败：{0}", e.Message); }
            SetStatus(n == 0 ? "没有改动，settings.ini 原样未动。" : "已保存 " + n + " 项，立刻生效。", false);
        }

        /// <summary>
        /// 离屏自检用：把某一项设成给定值并走一遍「保存」（和点按钮完全同一条路径）。
        /// ★存在的理由：真正会出错的不是 Settings.Set，而是"界面上的勾选框接了哪个
        ///   ini 键"这层接线 —— 接错了界面上一切正常、ini 里却写着别的项。
        /// </summary>
        internal bool SetRowForTest(string key, string value)
        {
            var r = _rows.FirstOrDefault(x => x.Key == key);
            if (r == null) return false;
            r.Write(value);
            SaveAll();
            // ★判据看**界面读值**，不看 ini 原文：写回默认值时 SaveAll 会把那一行重新
            //   注释掉，IniValue 于是变 null —— 那正是对的行为，不该判红。
            string now;
            try { now = r.Read(); } catch { return false; }
            return string.Equals(now, value, StringComparison.Ordinal);
        }

        /// <summary>
        /// 离屏自检用：这一项的勾选框现在是不是勾上的。
        /// ★反向项（`NO_XXX`）要验的正是"功能开着时框是勾上的"—— 只验 ini 读值验不到这一层，
        ///   而用户看到的只有框。
        /// </summary>
        internal bool RowCheckedForTest(string key)
        {
            var r = _rows.FirstOrDefault(x => x.Key == key);
            return r != null && r.Checked != null && r.Checked();
        }

        private void ResetDefaults()
        {
            if (InfoBox.Confirm("把这一页能调的项全部恢复成默认值？\r\n\r\n" +
                    "做法是把 settings.ini 里那几行重新注释掉（模板本身就是「注释掉 = 默认」），" +
                    "行上面的中文说明会留着。", "恢复默认") != DialogResult.OK) return;

            int n = 0;
            foreach (var r in _rows)
            {
                // 只处理"ini 里真的写过"的项：没写过的本来就是默认值，
                // 再去写一条注释行只会把 ini 越堆越长。
                if (Settings.IniValue(r.Key) == null) { r.Write(r.Def); continue; }
                if (Settings.Reset(r.Key, r.Def)) { r.Write(r.Def); r.Original = null; n++; }
            }
            try { _onSaved?.Invoke(); } catch { }
            SetStatus("已把 " + n + " 项恢复成默认值（settings.ini 里那几行已注释掉）。", false);
        }

        private void SetStatus(string text, bool error)
        {
            if (_status == null) return;
            _status.Text = text;
            _status.ForeColor = error ? UiTheme.Err : UiTheme.Ok;
            Log.Write("设置界面：{0}", text);
        }

        /// <summary>把页面套进一个 Fill 容器里（TableLayoutPanel 直接当页用会丢失滚动条）。</summary>
        private static Panel Wrap(TableLayoutPanel page)
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = UiTheme.Body, Visible = false };
            p.Controls.Add(page);
            return p;
        }
    }
}
