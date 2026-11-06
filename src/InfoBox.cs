// 托盘菜单的反馈框（模态）。
// ★不用 ShowBalloonTip：Win11 下用户在通知设置里关掉横幅、或图标被收进溢出区，
//   气泡一个像素都不显示，用户只会得出"这个菜单坏了"的结论。
// ★必须带一个屏幕外的 TopMost 属主窗口，否则 MessageBox 被压在其它窗口后面，等于没弹。

using System;
using System.Drawing;
using System.Windows.Forms;

namespace VbeComplete
{
    internal static class InfoBox
    {
        public static void Show(string title, string text)
        {
            try
            {
                using (var owner = new Form
                {
                    StartPosition = FormStartPosition.Manual,
                    Location = new Point(-32000, -32000),   // 屏幕外，纯粹当属主
                    Size = new Size(1, 1),
                    ShowInTaskbar = false,
                    FormBorderStyle = FormBorderStyle.None,
                    TopMost = true,
                })
                {
                    owner.Show();
                    owner.Activate();
                    MessageBox.Show(owner, text, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    owner.Hide();
                }
            }
            catch (Exception e) { Log.Write("提示框显示失败：{0}", e.Message); }
        }

        /// <summary>
        /// 确认框（是 / 否）。用它而不是裸 MessageBox 的理由和上面那条一样：
        /// 托盘程序没有主窗口，裸框会被压在当前窗口后面，用户以为"没反应"。
        /// </summary>
        public static DialogResult Confirm(string text, string title)
        {
            try
            {
                using (var owner = new Form
                {
                    StartPosition = FormStartPosition.Manual,
                    Location = new Point(-32000, -32000),
                    Size = new Size(1, 1),
                    ShowInTaskbar = false,
                    FormBorderStyle = FormBorderStyle.None,
                    TopMost = true,
                })
                {
                    owner.Show();
                    owner.Activate();
                    var r = MessageBox.Show(owner, text, title, MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    owner.Hide();
                    return r;
                }
            }
            catch (Exception e)
            {
                Log.Write("确认框显示失败：{0}", e.Message);
                return DialogResult.Cancel;
            }
        }
    }
}
