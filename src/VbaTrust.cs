// "信任对 VBA 工程对象模型的访问"（AccessVBOM）。
//
// ★WPS 默认是 0：此时 Application.VBE 直接抛 E_FAIL（0x80004005 "发生意外"），
//   看起来像"WPS 不支持对象模型"，其实只是这个开关没开。开了之后 VBE / VBProjects /
//   CodeModule / CodePane.SetSelection 在 WPS 上与 Office 完全一致（窗口类也一样）。
// ★WPS 读的是 Microsoft\Office\12.0\Excel\Security（它自称 "Microsoft Excel 12.0"），
//   但用户装的 Office 可能是别的版本 ⇒ 4 个版本号 × 3 个组件全写一遍，谁读都生效。
//   没有 Microsoft Office 的用户也没关系：多出来的注册表项是无害的。

using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace VbeComplete
{
    internal static class VbaTrust
    {
        private static readonly string[] Versions = { "12.0", "14.0", "15.0", "16.0" };
        private static readonly string[] Apps = { "Excel", "Word", "PowerPoint" };

        private const string KeyTemplate = @"Software\Microsoft\Office\{0}\{1}\Security";

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

        /// <summary>全部打开，返回实际改了几处（已经是 1 的不计）。</summary>
        public static int EnableAll()
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
                            if (k.GetValue("AccessVBOM") is int cur && cur == 1) continue;
                            k.SetValue("AccessVBOM", 1, RegistryValueKind.DWord);
                            changed++;
                        }
                    }
                    catch (Exception e)
                    {
                        Log.Write("写 AccessVBOM 失败（{0}/{1}）：{2}", v, a, e.Message);
                    }
                }
            return changed;
        }

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
