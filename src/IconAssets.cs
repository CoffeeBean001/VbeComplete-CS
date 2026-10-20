// 程序图标：从 exe 内嵌的 app.ico 里取，按需要的尺寸挑最合适的那一帧。
//
// ★为什么内嵌而不是读外部 文件：
//   1) 单文件发布时 exe 旁边没有 app.ico，读文件会变成"托盘一片空白"；
//   2) 托盘图标、任务栏图标、窗口图标必须是同一张图，用户才认得出；
//   3) .ico 里已经放了 16/20/24/32/40/48/64/128/256 九帧，`new Icon(stream, size)`
//      会自动挑最接近的一帧，不靠系统缩放（缩放出来的 16px 会糊）。
//
// ★`app.ico` 由 `tools/make-icon.py` 生成（深蓝圆角方块 + 白色 V），
//   改图标就改那个脚本再跑一遍，别手工塞位图。

using System;
using System.Drawing;
using System.IO;

namespace VbeComplete
{
    internal static class IconAssets
    {
        private const string ResourceName = "VbeComplete.app.ico";

        private static Icon _cached;
        private static int _cachedSize = -1;

        /// <summary>取指定像素尺寸的程序图标；失败时退回系统默认图标（绝不抛）。</summary>
        public static Icon Get(int size)
        {
            if (size <= 0) size = 32;
            if (_cached != null && _cachedSize == size) return _cached;
            try
            {
                var asm = typeof(IconAssets).Assembly;
                using (Stream s = asm.GetManifestResourceStream(ResourceName))
                {
                    if (s != null)
                    {
                        _cached = new Icon(s, new Size(size, size));
                        _cachedSize = size;
                        return _cached;
                    }
                }
                Log.Write("图标资源没找到（{0}），退回系统图标", ResourceName);
            }
            catch (Exception e) { Log.Write("图标加载失败：{0}", e.Message); }

            _cached = SystemIcons.Application;
            _cachedSize = size;
            return _cached;
        }

        /// <summary>托盘用的那一份（32px 足够，高 DPI 下系统会自己挑大图标）。</summary>
        public static Icon Tray() => Get(32);

        /// <summary>窗口标题栏 / 任务栏用（16px 是标题栏的实际尺寸）。</summary>
        public static Icon Window() => Get(32);
    }
}
