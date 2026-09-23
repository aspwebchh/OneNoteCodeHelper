using System;
using System.Runtime.InteropServices;
using System.Windows;

namespace OneNoteCodeHelper.Interop
{
    internal static class NativeMethods
    {
        private const uint MbSetForeground = 0x00010000;

        private const uint MbTopmost = 0x00040000;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBoxW(IntPtr owner, string text, string caption, uint type);

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
    }
}
