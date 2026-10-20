// 开机自动启动：往 HKCU\Software\Microsoft\Windows\CurrentVersion\Run 里登记一条。
//
// ★为什么用 HKCU 的 Run 键，而不是启动文件夹 / 计划任务：
//   1) 不需要管理员权限（UAC 一次都不弹），装在哪都能开；
//   2) 只有当前用户生效 —— 这是"我的编辑器插件"，不该影响别的账户；
//   3) 卸载/关掉就是删一个值，随时可逆，不碰系统计划任务库。
//
// ★"用户的选择"与"注册表当前状态"分开存：
//   注册表是**结果**（可以被别的优化软件删掉、被清理工具改掉），
//   `autostart.txt`（内容就一个 1/0）是**用户的意图**。
//   每次启动都按意图把注册表"修一遍"：
//     · 用户勾了 → 值不在 / 指向旧路径（exe 挪过地方）→ 重写
//     · 用户没勾 → 值还在（比如上次被强行写回）→ 删掉
//   这样换目录、被清理软件动过，都能自愈，不需要用户重新勾一次。
//
// ★默认是**开**（用户明确要求"开机自动启动"）；第一次运行没 marker 时按
//   settings.ini 的 AUTO_START 走（默认 1）。想关：托盘菜单取消勾选，
//   或者 settings.ini 里写 AUTO_START = 0 并把 autostart.txt 删掉。

using System;
using System.IO;
using Microsoft.Win32;

namespace VbeComplete
{
    internal static class AutoStart
    {
        private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string ValueName = "VbeComplete-CS";
        private const string StateFile = "autostart.txt";

        private static string StatePath => Path.Combine(Settings.Home, StateFile);

        /// <summary>要登记的完整命令行（带引号，防路径里有空格）。</summary>
        public static string CommandLine
        {
            get
            {
                string exe = ExePath();
                if (string.IsNullOrEmpty(exe)) return null;
                return "\"" + exe + "\"";
            }
        }

        /// <summary>
        /// 当前进程的真实 exe 路径。单文件发布时 `ProcessPath` 就是 `dist\VbeComplete.exe`；
        /// 极端情况下取不到，退回 `Application.ExecutablePath`（apphost 路径）。
        /// </summary>
        public static string ExePath()
        {
            try
            {
                string p = Environment.ProcessPath;
                if (!string.IsNullOrEmpty(p) &&
                    p.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                    return p;
            }
            catch { }
            try { return System.Windows.Forms.Application.ExecutablePath; } catch { }
            return null;
        }

        /// <summary>注册表里现在有没有这一条（顺带把值读出来，方便判断是不是指向旧路径）。</summary>
        public static bool IsRegistered(out string value)
        {
            value = null;
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RunKey, false))
                {
                    if (k == null) return false;
                    value = k.GetValue(ValueName) as string;
                    return !string.IsNullOrWhiteSpace(value);
                }
            }
            catch (Exception e)
            {
                Log.Write("读开机自启注册表失败：{0}", e.Message);
                return false;
            }
        }

        /// <summary>用户意图：读 marker，没有就按配置（默认开）。</summary>
        public static bool Desired
        {
            get
            {
                try
                {
                    if (File.Exists(StatePath))
                    {
                        string s = File.ReadAllText(StatePath).Trim();
                        if (s == "0") return false;
                        if (s == "1") return true;
                    }
                }
                catch { }
                return Settings.Flag("AUTO_START", true);
            }
        }

        /// <summary>用户从托盘菜单切换：先记住意图，再立刻落到注册表。</summary>
        public static bool SetEnabled(bool on)
        {
            try
            {
                File.WriteAllText(StatePath, on ? "1" : "0");
            }
            catch (Exception e) { Log.Write("写 {0} 失败（意图没记住，本次仍然生效）：{1}", StateFile, e.Message); }
            return Apply(on);
        }

        /// <summary>
        /// 启动时调用：按意图把注册表修正到"该有的样子"。
        /// 已一致就一个字节都不写（省得每次开机都碰注册表）。
        /// </summary>
        public static void Sync()
        {
            bool want = Desired;
            if (!want)
            {
                // 没勾：值还在就删掉（别的工具写回来的也一并清掉）
                if (IsRegistered(out _))
                {
                    if (Apply(false)) Log.Write("开机自启：已按你的设置关闭并清理注册表");
                }
                return;
            }

            string cmd = CommandLine;
            if (cmd == null) { Log.Write("开机自启：取不到 exe 路径，跳过"); return; }
            if (IsRegistered(out string cur) && string.Equals(cur, cmd, StringComparison.OrdinalIgnoreCase))
                return;                                  // 已经指向当前 exe，不用动

            if (Apply(true))
                Log.Write("开机自启：已登记 {0}（原值 {1}）", cmd, cur ?? "无");
        }

        private static bool Apply(bool on)
        {
            string cmd = CommandLine;
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(RunKey, true))
                {
                    if (k == null) { Log.Write("开机自启：打不开 Run 键"); return false; }
                    if (on)
                    {
                        if (cmd == null) return false;
                        k.SetValue(ValueName, cmd, RegistryValueKind.String);
                    }
                    else
                    {
                        // DeleteValue 对不存在的值会抛，先判断
                        if (k.GetValue(ValueName) != null) k.DeleteValue(ValueName, false);
                    }
                    return true;
                }
            }
            catch (Exception e)
            {
                Log.Write("开机自启：写注册表失败（可能被安全软件拦了）：{0}", e.Message);
                return false;
            }
        }

        /// <summary>给日志 / 托盘提示用的一句话状态。</summary>
        public static string Describe()
        {
            bool reg = IsRegistered(out string cur);
            if (Desired && reg) return "开机自启：已开启";
            if (!Desired && !reg) return "开机自启：已关闭";
            if (Desired && !reg) return "开机自启：应开启，但注册表里没有（可能被清理软件删了，下次启动会自愈）";
            return "开机自启：应关闭，但注册表里还有（下次启动会清掉）";
        }

        // ------------------------------------------------------------------
        // 自检用：先拍一张快照，测完原样还原 —— --startuptest 绝不许留下痕迹
        // ------------------------------------------------------------------

        internal sealed class State
        {
            internal bool HadValue;
            internal string Value;
            internal bool HadMarker;
            internal string Marker;
        }

        internal static State Capture()
        {
            var st = new State();
            IsRegistered(out string v);
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RunKey, false))
                {
                    if (k != null)
                    {
                        object raw = k.GetValue(ValueName);
                        if (raw != null) { st.HadValue = true; st.Value = raw as string; }
                    }
                }
            }
            catch { }
            try
            {
                if (File.Exists(StatePath))
                {
                    st.HadMarker = true;
                    st.Marker = File.ReadAllText(StatePath);
                }
            }
            catch { }
            return st;
        }

        internal static void Restore(State st)
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(RunKey, true))
                {
                    if (k != null)
                    {
                        if (st.HadValue) k.SetValue(ValueName, st.Value ?? "", RegistryValueKind.String);
                        else if (k.GetValue(ValueName) != null) k.DeleteValue(ValueName, false);
                    }
                }
            }
            catch (Exception e) { Log.Write("还原开机自启注册表失败：{0}", e.Message); }
            try
            {
                if (st.HadMarker) File.WriteAllText(StatePath, st.Marker ?? "");
                else if (File.Exists(StatePath)) File.Delete(StatePath);
            }
            catch (Exception e) { Log.Write("还原 {0} 失败：{1}", StateFile, e.Message); }
        }
    }
}
