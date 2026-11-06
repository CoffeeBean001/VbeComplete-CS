// "添加 / 编辑词条"小窗：一行一个批量粘，用户不用再开 DB Browser。
// ★函数库问"函数/类型 + 补括号"，关键字库问"补完跟空格"—— 两者语义不同
//   （`Dim ` 要空格、`MsgBox()` 要括号），不能混在一个勾选框里。
// ★只收集用户意图，落库在 WordBank；校验只挡"加了也用不上"的词（带空格、纯数字）。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace VbeComplete
{
    internal sealed class WordEditForm : Form
    {
        private readonly string _bank;
        private readonly string _oldWord;         // null = 新增
        private readonly bool _editMode;

        private readonly Font _uiFont;
        private readonly Font _boldFont;
        private readonly Font _monoFont;

        private TextBox _word;
        private ComboBox _kind;
        private CheckBox _chkSpace;               // 关键字：补完跟空格
        private CheckBox _chkParens;              // 函数：补完补括号
        private CheckBox _chkEnabled;
        private Label _hint;

        public string OldWord => _oldWord;
        /// <summary>确定后要写进库的词条（批量添加时可能有多条）。</summary>
        public List<WordEntry> Results { get; private set; } = new List<WordEntry>();

        private WordEditForm(string bank, WordEntry editing)
        {
            _bank = bank == WordBank.Functions ? WordBank.Functions : WordBank.Keywords;
            _editMode = editing != null;
            _oldWord = editing?.Word;
            bool kw = _bank == WordBank.Keywords;

            Text = (_editMode ? "编辑词条" : "添加词条") +
                   " — " + (kw ? "关键字库" : "函数与类型库");
            try { Icon = IconAssets.Window(); } catch { }
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            BackColor = UiTheme.Body;
            _uiFont = new Font("Microsoft YaHei UI", 9f);
            _boldFont = new Font("Microsoft YaHei UI", 10.5f, FontStyle.Bold);
            _monoFont = new Font("Consolas", 10f);
            Font = _uiFont;
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(470, _editMode ? 296 : 372);
            KeyPreview = true;

            BuildUi();
            if (_editMode) Fill(editing);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _uiFont?.Dispose(); _boldFont?.Dispose(); _monoFont?.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// 只建窗口、不弹模态 —— 给 --uitest 离屏渲染用（核对布局，不进用户流程）。
        /// </summary>
        public static WordEditForm Preview(string bank, WordEntry editing)
            => new WordEditForm(bank, editing);

        /// <summary>弹窗收集词条；用户取消返回 null，确定返回要写的条目（可能多条）。</summary>
        public static List<WordEntry> Ask(IWin32Window owner, string bank, WordEntry editing)
        {
            using (var f = new WordEditForm(bank, editing))
            {
                var r = owner != null ? f.ShowDialog(owner) : f.ShowDialog();
                return r == DialogResult.OK ? f.Results : null;
            }
        }

        // ------------------------------------------------------------------

        private void BuildUi()
        {
            bool kw = _bank == WordBank.Keywords;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                BackColor = UiTheme.Body,
                Padding = new Padding(16, 14, 16, 12),
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 标题
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 说明
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));  // 词条
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 选项
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 提示
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));   // 按钮

            var lblTitle = UiTheme.Label(_editMode ? "编辑词条" : "添加词条", UiTheme.Text, _boldFont);
            lblTitle.Margin = new Padding(0, 0, 0, 2);
            root.Controls.Add(lblTitle, 0, 0);

            var lblSub = UiTheme.Label(
                kw ? "关键字：出现在候选里的语言关键字（Dim / If / Then…），"
                   + "重复添加同一个词就是改它的设置。"
                   : "函数与类型：内建函数、类型名，或者你自己的常用函数；"
                     + "重复添加同一个词就是改它的设置。",
                UiTheme.Muted, _uiFont, true);
            lblSub.Margin = new Padding(0, 0, 0, 10);
            root.Controls.Add(lblSub, 0, 1);

            // ---- 词条 ----
            var wordArea = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = UiTheme.Body,
                Margin = new Padding(0),
                Padding = new Padding(0),
            };
            wordArea.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            wordArea.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            wordArea.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            var lblWord = UiTheme.Label(_editMode ? "词条" : "词条（一行一个，可一次粘多行）",
                UiTheme.Muted, _uiFont);
            lblWord.Margin = new Padding(0, 0, 0, 4);
            wordArea.Controls.Add(lblWord, 0, 0);

            _word = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = UiTheme.Card,
                ForeColor = UiTheme.Text,
                Font = _monoFont,
                Multiline = !_editMode,
                AcceptsReturn = !_editMode,
                ScrollBars = _editMode ? ScrollBars.None : ScrollBars.Vertical,
                WordWrap = false,
                Margin = new Padding(0),
            };
            wordArea.Controls.Add(_word, 0, 1);
            root.Controls.Add(wordArea, 0, 2);

            // ---- 选项 ----
            var opts = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                BackColor = UiTheme.Body,
                Margin = new Padding(0, 10, 0, 0),
            };

            if (kw)
            {
                _chkSpace = UiTheme.Check("补完后跟一个空格（`Dim ` 这样）", true);
                _chkSpace.Margin = new Padding(0, 0, 0, 2);
                opts.Controls.Add(_chkSpace);
            }
            else
            {
                var kindRow = new FlowLayoutPanel
                {
                    AutoSize = true, WrapContents = false, BackColor = UiTheme.Body,
                    Margin = new Padding(0, 0, 0, 2),
                };
                kindRow.Controls.Add(new Label
                {
                    Text = "种类", AutoSize = true, ForeColor = UiTheme.Muted,
                    BackColor = UiTheme.Body, Margin = new Padding(0, 6, 8, 0),
                });
                _kind = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Width = 250,
                    BackColor = UiTheme.Card,
                    ForeColor = UiTheme.Text,
                };
                _kind.Items.AddRange(new object[] { "函数（补一对括号）", "类型（类型名，不补括号）" });
                _kind.SelectedIndex = 0;
                _kind.SelectedIndexChanged += (s, e) => SyncKind();
                kindRow.Controls.Add(_kind);
                opts.Controls.Add(kindRow);

                _chkParens = UiTheme.Check("补完后补一对括号（`MsgBox()` 这样）", true);
                _chkParens.Margin = new Padding(0, 0, 0, 2);
                opts.Controls.Add(_chkParens);
            }

            _chkEnabled = UiTheme.Check("启用（取消勾选 = 词留着，但这段时间先不提示）", true);
            opts.Controls.Add(_chkEnabled);
            root.Controls.Add(opts, 0, 3);

            // ---- 提示行（校验错误写这里） ----
            _hint = UiTheme.Label("", UiTheme.Err, _uiFont, true);
            _hint.Dock = DockStyle.Fill;
            _hint.MinimumSize = new Size(0, 20);
            _hint.Margin = new Padding(0, 6, 0, 0);
            root.Controls.Add(_hint, 0, 4);

            // ---- 按钮 ----
            var bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                BackColor = UiTheme.Body,
                Margin = new Padding(0),
            };
            var ok = UiTheme.Primary("确定");
            ok.Click += (s, e) => Commit();
            var cancel = UiTheme.Plain("取消");
            cancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            bar.Controls.Add(ok);
            bar.Controls.Add(cancel);
            root.Controls.Add(bar, 0, 5);

            Controls.Add(root);
            CancelButton = cancel;
            if (_editMode) AcceptButton = ok;
            else
            {
                // 多行框里回车要能换行 ⇒ 不能挂 AcceptButton；用 Ctrl+Enter 提交。
                _word.KeyDown += (s, e) =>
                {
                    if (e.KeyCode == Keys.Enter && e.Control)
                    {
                        e.SuppressKeyPress = true;
                        Commit();
                    }
                };
            }
            SyncKind();
        }

        private void Fill(WordEntry e)
        {
            _word.Text = e.Word;
            _word.SelectionStart = _word.Text.Length;
            _chkEnabled.Checked = e.Enabled;
            if (_bank == WordBank.Keywords)
            {
                if (_chkSpace != null) _chkSpace.Checked = e.WantsSpace;
            }
            else
            {
                if (_kind != null) _kind.SelectedIndex = e.IsType ? 1 : 0;
                if (_chkParens != null) _chkParens.Checked = e.WantsParens;
            }
        }

        /// <summary>类型选成"类型名"时，括号那一项没意义（选了也不会补）。</summary>
        private void SyncKind()
        {
            if (_chkParens == null) return;
            bool isType = _kind != null && _kind.SelectedIndex == 1;
            _chkParens.Enabled = !isType;
            if (isType) _chkParens.Checked = false;
        }

        private void Commit()
        {
            _hint.ForeColor = UiTheme.Err;
            _hint.Text = "";

            var words = _word.Text.Replace("\r\n", "\n").Replace('\r', '\n')
                .Split('\n').Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
            if (words.Count == 0) { _hint.Text = "先填词条。"; _word.Focus(); return; }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var list = new List<WordEntry>();
            foreach (var w in words)
            {
                string why;
                if (!WordBank.IsValidWord(w, out why))
                {
                    _hint.Text = "「" + w + "」" + why;
                    _word.Focus();
                    return;
                }
                if (!seen.Add(w)) continue;                     // 同一批里的重复，静默去掉
                bool isType = _kind != null && _kind.SelectedIndex == 1;
                list.Add(new WordEntry
                {
                    Word = w,
                    Enabled = _chkEnabled.Checked,
                    WantsSpace = _chkSpace == null || _chkSpace.Checked,
                    Kind = isType ? "type" : "function",
                    WantsParens = !isType && _chkParens != null && _chkParens.Checked,
                });
            }

            Results = list;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
