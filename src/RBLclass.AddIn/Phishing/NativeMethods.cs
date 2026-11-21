using System;
using System.Runtime.InteropServices;
using System.Text;

namespace RBLclass.AddIn.Phishing
{
    /// <summary>
    /// The small Win32 surface used by the link warning (v2.13.0.0 D3): thread-scoped
    /// hooks on Outlook's own UI thread, window class and ancestor lookups, key state and
    /// the accessibility entry point that returns Word's object model for a document
    /// window. Every handle-sized field is an <see cref="IntPtr"/> so the same binary works
    /// in 32-bit and 64-bit Outlook. There is no global hook and nothing is injected.
    /// </summary>
    internal static class NativeMethods
    {
        public const int WH_GETMESSAGE = 3;
        public const int WH_MOUSE = 7;

        public const int WM_NULL = 0x0000;
        public const int WM_KEYDOWN = 0x0100;
        public const int WM_LBUTTONDOWN = 0x0201;
        public const int WM_LBUTTONUP = 0x0202;
        public const int WM_LBUTTONDBLCLK = 0x0203;
        public const int WM_RBUTTONDOWN = 0x0204;

        public const int PM_REMOVE = 1;
        public const int VK_RETURN = 0x0D;
        public const int VK_SHIFT = 0x10;
        public const int VK_CONTROL = 0x11;
        public const int VK_MENU = 0x12;
        public const uint GA_ROOT = 2;
        public const uint OBJID_NATIVEOM = 0xFFFFFFF0;

        public static readonly Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");

        public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEHOOKSTRUCT
        {
            public POINT pt;
            public IntPtr hwnd;
            public uint wHitTestCode;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public POINT pt;
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

        [DllImport("user32.dll")]
        public static extern short GetKeyState(int nVirtKey);

        [DllImport("user32.dll")]
        public static extern IntPtr GetFocus();

        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(
            IntPtr hwnd, uint dwObjectID, ref Guid riid,
            [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

        /// <summary>
        /// Standard OLE window interface, used only to compare an Outlook Explorer or
        /// Inspector with a window's root. The one hand-declared COM interface in the
        /// add-in: it is not an Office interface (the PIA rule does not apply).
        /// </summary>
        [ComImport]
        [Guid("00000114-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleWindow
        {
            [PreserveSig]
            int GetWindow(out IntPtr phwnd);

            [PreserveSig]
            int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
        }

        /// <summary>Window class name, or empty when it cannot be read.</summary>
        public static string ClassNameOf(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return string.Empty;
            var sb = new StringBuilder(64);
            return GetClassName(hwnd, sb, sb.Capacity) > 0 ? sb.ToString() : string.Empty;
        }
    }
}
