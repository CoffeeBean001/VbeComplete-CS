// 托盘菜单的反馈框。
//
// ★为什么不用气泡（NotifyIcon.ShowBalloonTip）：Win11 把老式气泡当"通知"投递，
//   用户在「设置 → 系统 → 通知」里关掉横幅、或者图标被收进托盘的溢出区，气泡就
//   **一个像素都不显示** —— 点了菜单，界面上毫无变化，用户只能得出"这个菜单坏了"
//   的结论（实测反馈："点了没有反应"）。凡是"点一下就该知道结果"的动作走模态框。
// ★必须带一个 TopMost 的属主窗口：托盘程序没有主窗口，裸 MessageBox.Show 弹出来
//   会被压在当前窗口**后面**，效果和没弹一样。属主放在屏幕外，只借它的 TopMost。
// ★模态期间消息泵被占住 —— 只有用户在点菜单时才会走到这里，此时没有打字，可接受。

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
