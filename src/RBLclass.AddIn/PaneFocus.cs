using System;
using System.Runtime.InteropServices;
using System.Text;

namespace RBLclass.AddIn
{
    /// <summary>
    /// Win32 focus checks for the task pane (v2.11.0.0 B3). WPF focus alone does
    /// not always take because Outlook holds the Win32 focus. The pane shares
    /// Outlook's UI thread, so <c>SetFocus</c> works without attaching input
    /// queues; no hook is involved.
    /// </summary>
    internal static class PaneFocus
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
        [DllImport("user32.dll")]
        private static extern IntPtr GetFocus();
        [DllImport("user32.dll", EntryPoint = "SetFocus")]
        private static extern IntPtr SetFocusNative(IntPtr hwnd);
        [DllImport("user32.dll")]
        private static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder name, int maxCount);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint GA_ROOT = 2;
        private const string MessageListClass = "OutlookGrid";

        /// <summary>
        /// True when the window hosting <paramref name="hwnd"/> is the foreground
        /// window, so taking focus cannot pull it away from a compose window that
        /// shares the UI thread.
        /// </summary>
        internal static bool IsRootForeground(IntPtr hwnd)
        {
            IntPtr root = GetAncestor(hwnd, GA_ROOT);
            return root != IntPtr.Zero && root == GetForegroundWindow();
        }

        internal static bool HasFocus(IntPtr hwnd)
        {
            IntPtr focused = GetFocus();
            return focused != IntPtr.Zero && (focused == hwnd || IsChild(hwnd, focused));
        }

        internal static void SetFocus(IntPtr hwnd)
        {
            SetFocusNative(hwnd);
        }

        /// <summary>
        /// Give the Win32 focus to Outlook's message list. After a programmatic
        /// folder change <c>Explorer.Activate</c> can leave no focused window at
        /// all, so the Delete key goes nowhere (#49). The class name is
        /// undocumented; when it is not found nothing happens and this returns
        /// false.
        /// </summary>
        internal static bool FocusMessageList(IntPtr hwnd)
        {
            IntPtr grid = IntPtr.Zero;
            EnumChildWindows(GetAncestor(hwnd, GA_ROOT), (child, _) =>
            {
                var name = new StringBuilder(32);
                GetClassName(child, name, name.Capacity);
                if (name.ToString() != MessageListClass || !IsWindowVisible(child)) return true;
                grid = child;
                return false;
            }, IntPtr.Zero);

            if (grid == IntPtr.Zero) return false;
            SetFocusNative(grid);
            return true;
        }
    }
}
