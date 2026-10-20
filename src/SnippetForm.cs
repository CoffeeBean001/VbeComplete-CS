// 代码片段窗（Ctrl+Q）。移植自 Python 版 snippet_ui.py，界面口径逐条对齐：
//   * 搜索框 —— 按标题模糊匹配（多词 AND，见 SnippetMatch）；
//   * 左边列表 = 全部片段标题，点一下右边就显示它的代码；
//   * 右边：标题框 + 代码框（Ctrl+Enter 直接插入、Tab 是缩进不是换焦点）；
//   * 底部：新增 / 保存修改 / 删除 / 关闭 / 插入到光标处；
//   * 开窗位置 = VBE 里光标处（下方放不下就翻到上方，再夹进屏幕）。
//
// ★与补全弹窗（PopupForm）的取向**刻意相反**：
//   补全窗是"只显示、不接收输入"的 NOACTIVATE 浮窗；这个窗要收键盘（搜索框里
//   要打字、代码框里要改代码），所以是普通窗口，会抢焦点、有标题栏、可拖动。
//   代价是"打开期间焦点不在 VBE"，全局钩子的焦点判据会把按键全部放行 —— 于是
//   窗内的 Esc / Ctrl+Q 必须由本窗自己绑（ProcessCmdKey）。
//
// 数据在 SnippetStore，插入动作由 App 注入（onInsert），本文件**不碰 COM**。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace VbeComplete
{
    internal sealed class SnippetForm : Form
    {
        // ---- 配色（与 Python 版 snippet_ui.py 顶部常量逐项对应） ----
        private static readonly Color CBg = ColorTranslator.FromHtml("#F7F6EF");
        private static readonly Color CPanel = ColorTranslator.FromHtml("#FFFFFF");
        private static readonly Color CFg = ColorTranslator.FromHtml("#202020");
        private static readonly Color CMuted = ColorTranslator.FromHtml("#6B6B6B");
        private static readonly Color CAccent = ColorTranslator.FromHtml("#2A5CAA");
        private static readonly Color CErr = ColorTranslator.FromHtml("#C00000");

        private const int DefaultW = 780;
        private const int DefaultH = 470;

        private readonly SnippetStore _store;
        private readonly Func<string, bool> _onInsert;     // (code) -> 是否插进去了
        public Action OnClosedCallback;                             // 收起之后（把焦点还给 VBE）

        private TextBox _search;
        private ListBox _list;
        private TextBox _title;
        private TextBox _code;
        private Label _status;
        private Button _btnInsert;

        private List<KeyValuePair<string, string>> _rows = new List<KeyValuePair<string, string>>();
        private List<SnippetHit> _view = new List<SnippetHit>();
        private string _shownTitle;
        private string _shownCode;
        private bool _bulk;                                 // 程序性改选中，别当用户点击

        public SnippetForm(SnippetStore store, Func<string, bool> onInsert)
        {
            _store = store ?? new SnippetStore();
            _onInsert = onInsert;

            Text = "代码片段";
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = CBg;
            Font = new Font("Microsoft YaHei UI", 9f);
            AutoScaleMode = AutoScaleMode.Font;
            MinimumSize = new Size(560, 300);
            Size = new Size(DefaultW, DefaultH);
            KeyPreview = true;

            BuildUi();
            FormClosing += (s, e) => { if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Close_(); } };
        }

        // ------------------------------------------------------------------
        // 建界面
        // ------------------------------------------------------------------

        private void BuildUi()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = CBg,
                Padding = new Padding(8, 8, 8, 8),
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // ---- 第一行：搜索 ----
            var bar = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 6),
            };
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            var lblSearch = new Label
            {
                Text = "搜索",
                AutoSize = true,
                ForeColor = CMuted,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 0, 6, 0),
            };
            _search = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = CPanel,
                ForeColor = CFg,
                Margin = new Padding(0, 0, 0, 0),
            };
            _search.TextChanged += (s, e) => Refilter(_shownTitle);
            // 搜索框里按 ↓ = 把焦点交给列表（并选中第一条）。Python 版（Tk）没有
            // 这一条，用户只能拿鼠标点列表；这里补上不影响任何既有行为 ——
            // TextBox 本来就不处理 ↓。
            _search.KeyDown += (s, e) =>
            {
                if (e.KeyCode != Keys.Down) return;
                if (_list.Items.Count == 0) return;
                if (_list.SelectedIndex < 0) _list.SelectedIndex = 0;
                _list.Focus();
                e.Handled = true;
                e.SuppressKeyPress = true;
            };

            bar.Controls.Add(lblSearch, 0, 0);
            bar.Controls.Add(_search, 1, 0);
            root.Controls.Add(bar, 0, 0);

            // ---- 第二行：左列表 + 右（标题 + 代码） ----
            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0),
            };
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 226f));
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            _list = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = CPanel,
                ForeColor = CFg,
                IntegralHeight = false,
                Margin = new Padding(0, 0, 6, 0),
            };
            _list.SelectedIndexChanged += (s, e) => OnPick();
            _list.KeyDown += ListKeyDown;
            _list.KeyPress += ListKeyPress;
            main.Controls.Add(_list, 0, 0);

            var right = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Margin = new Padding(0),
            };
            right.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            var lblTitle = new Label
            {
                Text = "标题",
                AutoSize = true,
                ForeColor = CMuted,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 0, 6, 4),
            };
            _title = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = CPanel,
                ForeColor = CFg,
                Margin = new Padding(0, 0, 0, 4),
            };
            right.Controls.Add(lblTitle, 0, 0);
            right.Controls.Add(_title, 1, 0);

            _code = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                AcceptsTab = true,          // Tab 是缩进，不是切焦点
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = CPanel,
                ForeColor = CFg,
                Font = new Font("Consolas", 10f),
                Margin = new Padding(0),
            };
            _code.KeyDown += CodeKeyDown;
            right.Controls.Add(_code, 0, 1);
            right.SetColumnSpan(_code, 2);
            main.Controls.Add(right, 1, 0);
            root.Controls.Add(main, 0, 1);

            // ---- 第三行：状态 + 按钮 ----
            var foot = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                AutoSize = true,
                Margin = new Padding(0, 6, 0, 0),
            };
            foot.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            foot.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _status = new Label
            {
                Text = "",
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = CMuted,
                Margin = new Padding(0),
            };
            foot.Controls.Add(_status, 0, 0);

            // ★FlowDirection=RightToLeft + 按"从右往左"的顺序加控件：
            //   第一个加进去的落在最右边 —— 与 Python 版 pack(side="right") 逐项一致。
            var btns = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                WrapContents = false,
                Margin = new Padding(0),
            };

            _btnInsert = new Button
            {
                Text = "插入到光标处",
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                BackColor = CAccent,
                ForeColor = Color.White,
                Padding = new Padding(8, 3, 8, 3),
                Margin = new Padding(6, 0, 0, 0),
            };
            _btnInsert.FlatAppearance.BorderSize = 0;
            _btnInsert.Click += (s, e) => DoInsert();
            btns.Controls.Add(_btnInsert);

            btns.Controls.Add(MakeButton("新增", DoAdd));
            btns.Controls.Add(MakeButton("保存修改", DoSave));
            btns.Controls.Add(MakeButton("删除", DoDelete));
            btns.Controls.Add(MakeButton("关闭", Close_));

            foot.Controls.Add(btns, 1, 0);
            root.Controls.Add(foot, 0, 2);

            Controls.Add(root);
        }

        private static Button MakeButton(string text, Action onClick)
        {
            var b = new Button
            {
                Text = text,
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = ColorTranslator.FromHtml("#202020"),
                Padding = new Padding(6, 3, 6, 3),
                Margin = new Padding(6, 0, 0, 0),
            };
            b.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#C8C8C8");
            b.Click += (s, e) => onClick();
            return b;
        }

        // ------------------------------------------------------------------
        // 显示 / 收起
        // ------------------------------------------------------------------

        /// <summary>
        /// 弹出并定位。`rect` = 光标屏幕坐标 (x, 上, 下)；拿不到就传 false。
        /// </summary>
        public void OpenAtCaret(bool hasRect, int x, int top, int bottom)
        {
            _bulk = true;
            try { _search.Text = ""; }
            finally { _bulk = false; }

            Reload();
            PlaceWindow(hasRect, x, top, bottom);

            if (!Visible) Show();
            TopMost = true;
            Activate();
            BringToFront();

            // 搜索框拿到焦点并全选 —— 打开就能直接打字过滤
            _search.Focus();
            _search.SelectAll();
        }

        private void PlaceWindow(bool hasRect, int x, int top, int bottom)
        {
            int w = Math.Max(MinimumSize.Width, Width);
            int h = Math.Max(MinimumSize.Height, Height);
            if (!hasRect)
            {
                // 拿不到光标：落在鼠标附近，别飞到屏幕角落去
                Native.POINT p;
                if (Native.GetCursorPos(out p)) { x = p.X; top = p.Y; bottom = p.Y + 16; }
                else { x = 120; top = 120; bottom = 136; }
            }

            Screen scr = Screen.FromPoint(new Point(x, Math.Max(0, bottom)));
            Rectangle wa = scr.WorkingArea;

            int py = bottom + 6;
            if (py + h > wa.Bottom) py = top - h - 6;      // 下方放不下 -> 翻到光标上方
            int px = x;
            if (px + w > wa.Right) px = wa.Right - w;
            if (px < wa.Left) px = wa.Left;
            if (py + h > wa.Bottom) py = wa.Bottom - h;
            if (py < wa.Top) py = wa.Top;

            Bounds = new Rectangle(px, py, w, h);
        }

        /// <summary>收起窗口，并把焦点交回 VBE（否则用户敲键盘会"没反应"）。</summary>
        public void Close_()
        {
            bool wasVisible = Visible;
            Hide();
            if (wasVisible)
            {
                try { OnClosedCallback?.Invoke(); } catch { }
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // 打开期间全局钩子不接管（焦点不在 VBE），所以这两个键必须自己处理
            if (keyData == Keys.Escape || keyData == (Keys.Control | Keys.Q))
            {
                // ★有未保存改动先问一句：直接关会让刚敲的代码静默消失
                //   （切到别的片段时反倒会问，关窗时却不问，说不过去）。
                if (IsDirty() && !Confirm("代码还没保存，确定关掉吗？")) return true;
                Close_();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // ------------------------------------------------------------------
        // 列表 / 内容
        // ------------------------------------------------------------------

        /// <summary>从库里重读全部片段，再按当前搜索词过滤。</summary>
        public void Reload(string keep = null)
        {
            try { _rows = _store.All(); }
            catch (Exception e) { _rows = new List<KeyValuePair<string, string>>(); Log.Write("读片段库失败：{0}", e.Message); }
            Refilter(keep ?? _shownTitle);
            // ★库打不开时 All() 只回一张空表，界面上就是"共 0 条"、列表全空 ——
            //   用户会以为片段全丢了。把真正的原因摆到状态栏上（覆盖掉条数提示）。
            string err = _store.LastError();
            if (!string.IsNullOrEmpty(err)) SetStatus("片段库打不开：" + err, true);
        }

        private void Refilter(string keep)
        {
            string q = _search == null ? "" : _search.Text;
            _view = SnippetMatch.Filter(_rows, q);

            _bulk = true;
            try
            {
                _list.BeginUpdate();
                _list.Items.Clear();
                foreach (var h in _view) _list.Items.Add(h.Title ?? "");
            }
            finally
            {
                _list.EndUpdate();
                _bulk = false;
            }

            int target = -1;
            if (!string.IsNullOrEmpty(keep))
            {
                for (int i = 0; i < _view.Count; i++)
                    if (string.Equals(_view[i].Title, keep, StringComparison.Ordinal)) { target = i; break; }
            }
            if (target < 0 && _view.Count > 0) target = 0;

            if (target >= 0)
            {
                SelectRow(target);
                // ★文本框里有未保存的改动时**不要**被搜索刷新覆盖掉
                if (!IsDirty()) ShowSnippet(_view[target].Title);
            }

            SetStatus(string.Format("共 {0} 条，命中 {1} 条", _rows.Count, _view.Count));
        }

        private void SelectRow(int index)
        {
            if (index < 0 || index >= _list.Items.Count) return;
            _bulk = true;
            try
            {
                _list.ClearSelected();
                _list.SelectedIndex = index;
            }
            finally { _bulk = false; }
        }

        private void OnPick()
        {
            if (_bulk) return;
            int i = _list.SelectedIndex;
            if (i < 0 || i >= _view.Count) return;
            string title = _view[i].Title;
            if (string.Equals(title, _shownTitle, StringComparison.Ordinal))
            {
                _title.Text = title ?? "";      // 标题框被手改过又点回自己：还原
                return;
            }
            if (IsDirty())
            {
                if (!Confirm("当前文本框里的内容改过但没保存，切换会丢掉这些修改。\n\n要继续切换吗？"))
                {
                    ReselectShown();
                    return;
                }
            }
            ShowSnippet(title);
        }

        private void ReselectShown()
        {
            for (int i = 0; i < _view.Count; i++)
            {
                if (string.Equals(_view[i].Title, _shownTitle, StringComparison.Ordinal))
                {
                    SelectRow(i);
                    return;
                }
            }
        }

        /// <summary>文本框内容是否与"当前展示的那条"不一致（纯比较，不靠事件记账）。</summary>
        private bool IsDirty()
        {
            if (_shownTitle == null) return CodeText().Trim().Length > 0;
            return !string.Equals(CodeText(), _shownCode ?? "", StringComparison.Ordinal);
        }

        private string CodeText() => _code == null ? "" : (_code.Text ?? "");

        /// <summary>
        /// 库存内容一律是 LF（Python 版 Tk 存进去的就是 LF），而 WinForms 的多行
        /// TextBox 只认 CRLF —— 不换算的话整段代码会挤成一行（真机就踩到了）。
        /// 存/插之前再换回 LF，两边读到的字节形态保持一致。
        /// </summary>
        private static string ToCrlf(string s) =>
            (s ?? "").Replace("\r\n", "\n").Replace("\n", "\r\n");

        private static string ToLf(string s) =>
            (s ?? "").Replace("\r\n", "\n").Replace('\r', '\n');

        private void ShowSnippet(string title)
        {
            string code = ToCrlf(_store.Get(title));
            _shownTitle = title;
            _shownCode = code;
            _code.Text = code;
            _title.Text = title ?? "";
            _code.SelectionStart = 0;
            _code.SelectionLength = 0;
        }

        private void SetStatus(string msg, bool err = false)
        {
            if (_status == null) return;
            _status.Text = msg ?? "";
            _status.ForeColor = err ? CErr : CMuted;
        }

        private bool Confirm(string text)
        {
            return MessageBox.Show(this, text, "代码片段", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes;
        }

        private void ListKeyPress(object sender, KeyPressEventArgs e)
        {
            // 在列表框里直接打字 -> 把这一下送进搜索框（别让 Listbox 自己跳字）
            if (!char.IsControl(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar))
            {
                _search.Focus();
                _search.Text += e.KeyChar;
                _search.SelectionStart = _search.Text.Length;
                e.Handled = true;
            }
        }

        private void ListKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !e.Shift && !e.Control)
            {
                DoInsert();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void CodeKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                DoInsert();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }
            if (e.KeyCode == Keys.Tab && !e.Control && !e.Alt)
            {
                // 文本框里的 Tab 应该是"缩进"，不是"换焦点"
                int at = _code.SelectionStart;
                _code.SelectedText = "    ";
                _code.SelectionStart = at + 4;
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        // ------------------------------------------------------------------
        // 四个动作
        // ------------------------------------------------------------------

        private void DoAdd()
        {
            string title = (_title.Text ?? "").Trim();
            string code = CodeText();
            if (title.Length == 0)
            {
                SetStatus("先填标题，再点「新增」", true);
                _title.Focus();
                return;
            }
            if (code.Trim().Length == 0)
            {
                SetStatus("代码是空的，没法存", true);
                return;
            }
            if (_store.Add(title, ToLf(code)))
            {
                Reload(title);
                // 刚存进去的内容与文本框里一模一样 —— 显式 ShowSnippet 一次，让
                // "当前展示项"认到这条新记录上（否则 IsDirty 会一直认为有改动，
                // 之后搜索 / 切换都不敢动文本框）。
                ShowSnippet(title);
                SetStatus("已新增「" + title + "」");
            }
            else
            {
                SetStatus("标题「" + title + "」已存在 —— 想改它请用「保存修改」", true);
            }
        }

        private void DoSave()
        {
            string cur = _shownTitle;
            if (string.IsNullOrEmpty(cur))
            {
                SetStatus("先在左边选一条要修改的片段", true);
                return;
            }
            string title = (_title.Text ?? "").Trim();
            string code = CodeText();
            if (title.Length == 0)
            {
                SetStatus("标题不能为空", true);
                return;
            }
            if (!string.Equals(title, cur, StringComparison.Ordinal))
            {
                if (_store.Exists(title))
                {
                    SetStatus("标题「" + title + "」已被别的片段占用", true);
                    return;
                }
                if (!_store.Rename(cur, title))
                {
                    SetStatus("改标题失败", true);
                    return;
                }
                _store.Update(title, ToLf(code));
                Reload(title);
                ShowSnippet(title);          // 认到新标题上（见 DoAdd 里的说明）
                SetStatus("已保存（标题改为「" + title + "」）");
                return;
            }
            if (_store.Update(title, ToLf(code)))
            {
                Reload(title);
                ShowSnippet(title);
                SetStatus("已保存「" + title + "」");
            }
            else
            {
                SetStatus("保存失败（这条片段可能已被删掉）", true);
            }
        }

        private void DoDelete()
        {
            string cur = _shownTitle;
            if (string.IsNullOrEmpty(cur))
            {
                SetStatus("先在左边选一条要删除的片段", true);
                return;
            }
            if (!Confirm("确定删除片段「" + cur + "」？\n\n此操作不可撤销。")) return;

            if (_store.Delete(cur))
            {
                _shownTitle = null;
                _shownCode = null;
                _code.Text = "";
                _title.Text = "";
                Reload(null);
                SetStatus("已删除「" + cur + "」");
            }
            else
            {
                SetStatus("删除失败", true);
            }
        }

        private void DoInsert()
        {
            string code = CodeText();
            if (code.Trim().Length == 0)
            {
                SetStatus("没有可插入的代码", true);
                return;
            }
            bool ok = false;
            try { ok = _onInsert != null && _onInsert(code); }
            catch (Exception e) { Log.Write("插入片段异常：{0}", e.Message); }
            if (!ok)
            {
                SetStatus("插入失败 —— 光标是不是已经离开代码窗格了？", true);
                return;
            }
            Close_();                       // 收起（OnClosed 里把焦点交回 VBE）
        }

        // ------------------------------------------------------------------
        // 给 --uitest 用：把窗子离屏画成一张图，用来核对布局配色
        // ------------------------------------------------------------------
        public void PreviewOnly(List<KeyValuePair<string, string>> rows)
        {
            _rows = rows ?? new List<KeyValuePair<string, string>>();
            _search.Text = "";
            Refilter(null);
        }
    }
}
