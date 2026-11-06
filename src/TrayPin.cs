// 让托盘图标常驻（不被 Win11 折叠进 ^ 溢出区）。
// ★Win11 把每个图标的状态记在 HKCU\Control Panel\NotifyIconSettings\<id>：
//   按 ExecutablePath 精确找到自己那条，只把 IsPromoted 置 1，绝不碰别人的。
// ★这条记录在图标第一次出现时才创建，找不到就安静返回。

using System;
using Microsoft.Win32;

namespace VbeComplete
{
    internal static class TrayPin
    {
        private const string Root = @"Control Panel\NotifyIconSettings";

        /// <summary>
        /// 尽力把自己的托盘图标设为"始终显示"。只写日志，不抛异常。
        /// </summary>
        public static void EnsureVisible()
        {
            string exe = AutoStart.ExePath();
            if (string.IsNullOrEmpty(exe)) return;

            int seen = 0, changed = 0;
            try
            {
                using (var root = Registry.CurrentUser.OpenSubKey(Root, true))
                {
                    if (root == null) return;          // 老系统 / 精简版没有这个键
                    foreach (string id in root.GetSubKeyNames())
                    {
                        try
                        {
                            using (var k = root.OpenSubKey(id, true))
                            {
                                if (k == null) continue;
                                string p = k.GetValue("ExecutablePath") as string;
                                if (string.IsNullOrEmpty(p)) continue;
                                if (!string.Equals(p, exe, StringComparison.OrdinalIgnoreCase)) continue;

                                seen++;
                                object raw = k.GetValue("IsPromoted");
                                int cur = raw is int i ? i : (raw is long l ? (int)l : 0);
                                if (cur != 1)
                                {
                                    k.SetValue("IsPromoted", 1, RegistryValueKind.DWord);
                                    changed++;
                                }
                            }
                        }
                        catch (Exception e) { Log.Write("托盘常驻：处理 {0} 失败：{1}", id, e.Message); }
                    }
                }
                if (changed > 0)
                    Log.Write("托盘常驻：已把本程序图标设为始终显示（{0} 条，重启资源管理器/重新登录后生效）",
                        changed);
                else if (seen > 0)
                    Log.Write("托盘常驻：已是最新（找到 {0} 条，无需改动）", seen);
                else
                    Log.Write("托盘常驻：还没找到本程序的托盘记录（首次运行属正常，下次启动会再试）");
            }
            catch (Exception e) { Log.Write("托盘常驻：写入失败（不影响使用）：{0}", e.Message); }
        }
    }
}
