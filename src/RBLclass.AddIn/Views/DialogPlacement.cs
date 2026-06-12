using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Interop;
using Serilog;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// Parents a WPF dialog to the host Outlook window so it opens on the same
    /// monitor as Outlook (paired with <c>WindowStartupLocation="CenterOwner"</c>)
    /// instead of always on the primary screen. The add-in runs inside the
    /// Outlook process, so the process main-window handle is Outlook's main
    /// explorer; owning the dialog by it also keeps the dialog above Outlook and
    /// minimised/restored with it.
    /// </summary>
    internal static class DialogPlacement
    {
        public static void OwnByOutlook(Window window)
        {
            if (window == null) return;
            try
            {
                IntPtr hwnd = Process.GetCurrentProcess().MainWindowHandle;
                if (hwnd != IntPtr.Zero)
                    new WindowInteropHelper(window).Owner = hwnd;
            }
            catch (Exception ex)
            {
                // Non-fatal: without an owner the dialog falls back to centring on
                // the primary screen, which is merely the prior behaviour.
                try { Log.Debug(ex, "Could not parent dialog to the Outlook window."); } catch { }
            }
        }
    }
}
