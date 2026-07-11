using System;
using System.Runtime.InteropServices;

namespace RBLclass.AddIn
{
    /// <summary>
    /// Measures how far Outlook's left navigation icon rail (the "app bar" /
    /// Hub Bar added to classic Outlook in Version 2211) overlaps the task-pane
    /// host window. Office's docking manager gives a Top/Bottom-docked Custom
    /// Task Pane the full client width, but the rail is a separate child window
    /// painted over its leftmost ~48px, and the classic CTP API exposes no way
    /// to move the pane out from under it. The workaround: measure the overlap
    /// and pad the host content right by that amount, so nothing useful renders
    /// in the hidden strip.
    ///
    /// The rail is detected geometrically rather than by class name (observed:
    /// a NetUINativeHWNDHost plus a mirroring MsoCommandBarDock strip, both
    /// undocumented): any visible direct child of the Outlook Explorer window
    /// that is a narrow, tall strip covering the host's left edge. When the
    /// rail is hidden ("Show apps in Outlook" unchecked), the pane is docked
    /// Left/Right, or the pane floats, nothing matches and the inset is 0.
    /// </summary>
    internal static class NavRailInset
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);

        private const uint GA_ROOT = 2;

        // A rail strip is at most this wide (48px at 100% scaling; generous
        // headroom for high-DPI monitors) - anything wider is a content pane.
        private const int MaxRailWidthPixels = 150;

        /// <summary>
        /// Pixels of the host window's left edge covered by a rail-like strip,
        /// clamped to [0, MaxRailWidthPixels]. All values are physical pixels,
        /// the same space WinForms control padding uses in this process.
        /// </summary>
        public static int MeasureLeftOverlapPixels(IntPtr hostHwnd)
        {
            if (hostHwnd == IntPtr.Zero) return 0;
            RECT host;
            if (!GetWindowRect(hostHwnd, out host) || host.Right <= host.Left) return 0;

            IntPtr root = GetAncestor(hostHwnd, GA_ROOT);
            if (root == IntPtr.Zero || root == hostHwnd) return 0;
            RECT rootRect;
            if (!GetWindowRect(root, out rootRect)) return 0;
            int minRailHeight = (rootRect.Bottom - rootRect.Top) / 3;

            int inset = 0;
            EnumChildWindows(root, (h, _) =>
            {
                // Direct children of the Explorer frame only; never the host
                // itself or one of its own ancestors/descendants.
                if (GetParent(h) != root || h == hostHwnd ||
                    IsChild(h, hostHwnd) || IsChild(hostHwnd, h))
                    return true;
                if (!IsWindowVisible(h)) return true;

                RECT r;
                if (!GetWindowRect(h, out r)) return true;
                int width = r.Right - r.Left;
                int height = r.Bottom - r.Top;

                bool railLike =
                    width > 0 && width <= MaxRailWidthPixels &&
                    height >= minRailHeight &&
                    r.Left <= host.Left &&      // starts at or left of the host
                    r.Right > host.Left &&      // actually covers its left edge
                    r.Top < host.Bottom && r.Bottom > host.Top; // vertical overlap

                if (railLike && r.Right - host.Left > inset)
                    inset = r.Right - host.Left;
                return true;
            }, IntPtr.Zero);

            return Math.Min(inset, MaxRailWidthPixels);
        }
    }
}
