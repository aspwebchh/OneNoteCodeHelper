using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows;

namespace OneNoteCodeHelper.Interop
{
    internal static class NativeMethods
    {
        private const uint MbSetForeground = 0x00010000;

        private const uint MbTopmost = 0x00040000;

        private const uint GaRoot = 2;

        private const uint MonitorDefaultToNearest = 2;

        private const uint SwpNoSize = 0x0001;

        private const uint SwpNoZOrder = 0x0004;

        private const uint SwpNoActivate = 0x0010;

        private const int StreamSeekSet = 0;

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MonitorInfo
        {
            public int Size;
            public Rect Monitor;
            public Rect Work;
            public uint Flags;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBoxW(IntPtr owner, string text, string caption, uint type);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsIconic(IntPtr hwnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPos(IntPtr hwnd, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);

        [DllImport("ole32.dll")]
        private static extern int CreateStreamOnHGlobal(IntPtr global, [MarshalAs(UnmanagedType.Bool)] bool deleteOnRelease,
            out IStream stream);

        /// <summary>
        /// 弹一个以 owner 为属主的提示框，阻塞到用户点掉为止。
        ///
        /// 用 Win32 的 MessageBox 而不是 WPF 的，是因为 WPF 版只接受 WPF Window 做属主，
        /// 而这里的属主是 OneNote 进程里的窗口。插件跑在 dllhost 里，不是前台进程，
        /// 没有属主的框很可能被压在 OneNote 后面；认 OneNote 窗口做属主，框就一定浮在它上面。
        /// </summary>
        internal static void ShowMessageBox(IntPtr owner, string text, string caption, MessageBoxImage icon)
        {
            // WPF 的 MessageBoxImage 取值和 Win32 的 MB_ICON* 一致，可以直接用。
            var type = (uint)icon | MbSetForeground;

            if (owner != IntPtr.Zero && MessageBoxW(owner, text, caption, type) != 0)
            {
                return;
            }

            // 没有属主，或者属主窗口已经没了（这时 MessageBox 直接返回 0）：退回置顶弹出。
            MessageBoxW(IntPtr.Zero, text, caption, type | MbTopmost);
        }

        /// <summary>
        /// 把 window 挪到 owner 当前所在位置的正中，并保证不超出 owner 所在显示器的工作区。
        ///
        /// 不用 WPF 的 CenterOwner：属主不是 WPF 窗口时，WPF 取的是 GetWindowPlacement 的
        /// rcNormalPosition，也就是属主「还原后」的位置。OneNote 通常是最大化的，
        /// 窗口就会居中到那个看不见的还原矩形上，看起来是偏的。
        ///
        /// 全程用设备像素：两个窗口的坐标都由本进程查询，处在同一个 DPI 坐标空间里，不用换算。
        /// 任何一步拿不到数据就什么都不做，保留 WPF 已经算好的位置。
        /// </summary>
        internal static void CenterOver(IntPtr window, IntPtr owner)
        {
            if (window == IntPtr.Zero || owner == IntPtr.Zero)
            {
                return;
            }

            // OneNote 给的句柄万一不是顶层框架窗口，按它所在的顶层窗口算。
            var root = GetAncestor(owner, GaRoot);
            if (root != IntPtr.Zero)
            {
                owner = root;
            }

            var monitor = MonitorFromWindow(owner, MonitorDefaultToNearest);
            var info = new MonitorInfo { Size = Marshal.SizeOf(typeof(MonitorInfo)) };
            if (monitor == IntPtr.Zero || !GetMonitorInfo(monitor, ref info))
            {
                return;
            }

            var work = info.Work;

            // 属主最小化时 GetWindowRect 返回的是屏幕外的 (-32000, -32000)，这时改为居中到工作区。
            Rect target;
            if (IsIconic(owner) || !GetWindowRect(owner, out target))
            {
                target = work;
            }

            if (!GetWindowRect(window, out var self))
            {
                return;
            }

            var width = self.Right - self.Left;
            var height = self.Bottom - self.Top;
            var x = Clamp(target.Left + (target.Right - target.Left - width) / 2, work.Left, work.Right - width);
            var y = Clamp(target.Top + (target.Bottom - target.Top - height) / 2, work.Top, work.Bottom - height);

            SetWindowPos(window, IntPtr.Zero, x, y, 0, 0, SwpNoSize | SwpNoZOrder | SwpNoActivate);
        }

        /// <summary>
        /// 把字节包成一个 COM 流，读指针在开头。
        ///
        /// 功能区的 loadImage 回调在 OneNote 里必须返回 IStream：插件跑在 dllhost 里，
        /// IPictureDisp 里装的是本进程的 GDI 句柄，过不了进程边界。
        /// 用系统自带的 HGlobal 流而不是自己实现 IStream，跨进程封送由系统负责。
        /// </summary>
        internal static IStream CreateStream(byte[] data)
        {
            Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out var stream));
            stream.Write(data, data.Length, IntPtr.Zero);
            stream.Seek(0, StreamSeekSet, IntPtr.Zero);
            return stream;
        }

        /// <summary>窗口比工作区还大时贴着左 / 上边，否则保证整个落在工作区里。</summary>
        private static int Clamp(int value, int min, int max)
        {
            return value > max ? Math.Max(min, max) : Math.Max(min, value);
        }
    }
}
