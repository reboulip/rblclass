using System;
using System.Runtime.InteropServices;
using RBLclass.Core;
using Serilog;
using OutlookOM = Microsoft.Office.Interop.Outlook;

namespace RBLclass.AddIn.Phishing
{
    /// <summary>
    /// From a clicked Word document window to the Outlook mail behind it and its sender
    /// (v2.13.0.0 D3). The window's root is compared with the Explorer's and each
    /// Inspector's window through the standard OLE <c>IOleWindow</c>. Only envelope
    /// properties are read (never the body). Every COM object is released except an
    /// item that Outlook owns. Anything unresolvable returns <see cref="SenderInfo.Unknown"/>
    /// (the warning still shows, without the sender line).
    /// </summary>
    internal sealed class LinkMailResolver
    {
        private readonly OutlookOM.Application _app;
        private readonly IMailStore _store;

        public LinkMailResolver(OutlookOM.Application app, IMailStore store)
        {
            _app = app;
            _store = store;
        }

        public SenderInfo ResolveSender(IntPtr wwgHwnd)
        {
            try
            {
                IntPtr root = NativeMethods.GetAncestor(wwgHwnd, NativeMethods.GA_ROOT);
                if (root == IntPtr.Zero || _app == null) return SenderInfo.Unknown;

                // Explorer: the reading pane (or an inline reply) shows the selected mail.
                OutlookOM.Explorer explorer = null;
                try { explorer = _app.ActiveExplorer(); } catch { }
                if (explorer != null)
                {
                    try
                    {
                        if (WindowOf(explorer) == root)
                            return SenderOfSelection(explorer);
                    }
                    finally { Release(explorer); }
                }

                // Inspector: an opened mail, or a reply being composed.
                OutlookOM.Inspectors inspectors = null;
                try { inspectors = _app.Inspectors; } catch { }
                if (inspectors == null) return SenderInfo.Unknown;
                try
                {
                    int count = inspectors.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        OutlookOM.Inspector inspector = null;
                        try { inspector = inspectors[i]; } catch { continue; }
                        try
                        {
                            if (WindowOf(inspector) != root) continue;
                            return SenderOfInspector(inspector);
                        }
                        finally { Release(inspector); }
                    }
                }
                finally { Release(inspectors); }
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Link warning: resolving the mail failed.");
            }
            return SenderInfo.Unknown;
        }

        private SenderInfo SenderOfSelection(OutlookOM.Explorer explorer)
        {
            OutlookOM.Selection selection = null;
            try
            {
                selection = explorer.Selection;
                if (selection == null || selection.Count != 1) return SenderInfo.Unknown;
                object item = selection[1];
                try { return _store.GetSenderInfo(item); }
                finally { Release(item); }
            }
            finally { Release(selection); }
        }

        private SenderInfo SenderOfInspector(OutlookOM.Inspector inspector)
        {
            object item = null;
            try
            {
                item = inspector.CurrentItem;
                var mail = item as OutlookOM.MailItem;
                // A draft (a reply being composed) has no external sender to show.
                if (mail != null && !mail.Sent) return SenderInfo.Unknown;
                return _store.GetSenderInfo(item);
            }
            finally { Release(item); }
        }

        private static IntPtr WindowOf(object outlookObject)
        {
            var oleWindow = outlookObject as NativeMethods.IOleWindow;
            if (oleWindow == null) return IntPtr.Zero;
            IntPtr hwnd;
            return oleWindow.GetWindow(out hwnd) == 0 ? hwnd : IntPtr.Zero;
        }

        private static void Release(object com)
        {
            if (com == null) return;
            try { if (Marshal.IsComObject(com)) Marshal.ReleaseComObject(com); } catch { }
        }
    }
}
