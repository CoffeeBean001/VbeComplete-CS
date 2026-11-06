// "信任对 VBA 工程对象模型的访问"（AccessVBOM）。
// ★WPS 默认是 0：此时 Application.VBE 直接 E_FAIL，看着像"WPS 不支持对象模型"，
//   其实只是这个开关没开。
// ★WPS 读 Microsoft\Office\12.0\Excel\Security，但用户装的 Office 版本不定 ⇒
//   4 个版本号 × 3 个组件全写一遍，谁读都生效。
// ★这是用户的开关：他关掉之后绝不能在背后自动写回 1（trust-off.txt 记着他的选择）。

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Win32;

namespace VbeComplete
{
    internal static class VbaTrust
    {
        private static readonly string[] Versions = { "12.0", "14.0", "15.0", "16.0" };
        private static readonly string[] Apps = { "Excel", "Word", "PowerPoint" };

        private const string KeyTemplate = @"Software\Microsoft\Office\{0}\{1}\Security";

        /// <summary>用户主动关掉过的标记（内容无意义，存在即"他关过"）。</summary>
        private const string OffMarker = "trust-off.txt";
        private static string OffPath => Path.Combine(Settings.Home, OffMarker);

        private static object Read(string ver, string app)
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(string.Format(KeyTemplate, ver, app)))
                    return k?.GetValue("AccessVBOM");
            }
            catch { return null; }
        }

        /// <summary>
        /// 有任何一个 AccessVBOM = 0 就算被禁。值不存在按"没禁"处理 ——
        /// Office 的默认行为是允许，只有用户/软件显式写过 0 才是关的。
        /// </summary>
        public static bool IsBlocked()
        {
            foreach (var v in Versions)
                foreach (var a in Apps)
                    if (Read(v, a) is int i && i == 0) return true;
            return false;
        }

        /// <summary>
        /// 12 处**全是 1** 才算开着（缺值按"没开"算 —— Office 的出厂默认就是关的）。
        /// ★和 IsBlocked 不是一回事：那个只认"显式写过 0"，用来决定要不要自动动手；
        ///   这个用来给用户看真实状态（托盘菜单的勾）。
        /// </summary>
        public static bool AllEnabled()
        {
            foreach (var v in Versions)
                foreach (var a in Apps)
                    if (!(Read(v, a) is int i && i == 1)) return false;
            return true;
        }

        /// <summary>逐版本一行的当前值，给提示框看（用户能自己核对）。</summary>
        public static string Detail()
        {
            var sb = new StringBuilder();
            foreach (var v in Versions)
            {
                sb.Append("  " + v + "   ");
                foreach (var a in Apps)
                {
                    object val = Read(v, a);
                    sb.Append(a + "=" + (val == null ? "未设置" : val.ToString()) + "   ");
                }
                if (v != Versions[Versions.Length - 1]) sb.Append("\r\n");
            }
            return sb.ToString();
        }

        /// <summary>全部打开，返回实际改了几处（已经是 1 的不计）。</summary>
        public static int EnableAll() => Write(1);

        /// <summary>全部关掉（写 0），返回实际改了几处。</summary>
        public static int DisableAll() => Write(0);

        private static int Write(int value)
        {
            int changed = 0;
            foreach (var v in Versions)
                foreach (var a in Apps)
                {
                    try
                    {
                        using (var k = Registry.CurrentUser.CreateSubKey(string.Format(KeyTemplate, v, a)))
                        {
                            if (k == null) continue;
                            if (k.GetValue("AccessVBOM") is int cur && cur == value) continue;
                            k.SetValue("AccessVBOM", value, RegistryValueKind.DWord);
                            changed++;
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Write("写 AccessVBOM 失败（{0}/{1}）：{2}", v, a, e.Message);
                    }
                }
            Remember(value == 1);
            return changed;
        }

        /// <summary>
        /// 记录"这是用户自己的选择"：关掉就留标记，打开就删掉。
        /// ★标记只用来挡住**后面**的自动帮忙（见 ExplainVbeDenied），不参与状态判断 ——
        ///   状态永远以注册表为准，用户在 Office 选项里手工改过也认。
        /// </summary>
        private static void Remember(bool on)
        {
            try
            {
                if (on)
                {
                    if (File.Exists(OffPath)) File.Delete(OffPath);
                }
                else File.WriteAllText(OffPath, "0");
            }
            catch (Exception e) { Log.Write("写 {0} 失败：{1}", OffMarker, e.Message); }
        }

        /// <summary>
        /// 用户从托盘**主动关掉**过，而且现在确实是关的。
        /// ★"现在确实是关的"这一半很关键：他后来在 Office 界面里自己开了，旧标记就该作废，
        ///   否则我们会一直拒绝自动帮忙（明明是他开的）。
        /// </summary>
        public static bool UserDisabled => File.Exists(OffPath) && !AllEnabled();

        /// <summary>日志 / 气泡里显示当前值。</summary>
        public static string Describe()
        {
            var sb = new StringBuilder();
            foreach (var v in Versions)
                foreach (var a in Apps)
                {
                    object val = Read(v, a);
                    if (val != null) sb.Append(v + "/" + a + "=" + val + " ");
                }
            return sb.Length == 0 ? "(未设置)" : sb.ToString().TrimEnd();
        }
    }
}
