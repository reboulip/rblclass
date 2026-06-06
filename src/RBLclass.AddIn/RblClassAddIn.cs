using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Office = Microsoft.Office.Core;
// Aliased as OutlookOM (not "Outlook") because the RBLclass.Outlook.Adapter
// project introduces an RBLclass.Outlook namespace, which would shadow a plain
// "Outlook" alias from inside the RBLclass.AddIn namespace.
using OutlookOM = Microsoft.Office.Interop.Outlook;

namespace RBLclass.AddIn
{
    /// <summary>
    /// RBLclass Outlook COM add-in shell. Thin IDTExtensibility2 +
    /// IRibbonExtensibility host; all business logic lives in RBLclass.Core
    /// and all Outlook OM access in RBLclass.Outlook.Adapter.
    /// </summary>
    [ComVisible(true)]
    [Guid(ClsidString)]
    [ProgId(ProgIdString)]
    // AutoDispatch (not None) so Office can resolve the ribbon onAction
    // callbacks (e.g. OnAboutClick) via IDispatch::GetIDsOfNames on the class
    // itself. With None, only members of explicitly-implemented interfaces are
    // findable and the ribbon callbacks would fail at click time. See
    // CLAUDE.md "COM interop interface declarations".
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class RblClassAddIn : IDTExtensibility2, Office.IRibbonExtensibility
    {
        // Kept in sync with release.config.json (Clsid / ProgId) - the source
        // of truth for the installer's HKCU COM registration.
        public const string ClsidString = "43808654-547E-4222-9BE3-7FA4B781FA44";
        public const string ProgIdString = "RBLclass.AddIn";

        private OutlookOM.Application _outlookApp;

        public void OnConnection(object application,
                                 ext_ConnectMode connectMode,
                                 object addInInst,
                                 ref Array custom)
        {
            try
            {
                _outlookApp = (OutlookOM.Application)application;
            }
            catch (Exception ex)
            {
                ShowError("OnConnection failed", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                if (_outlookApp != null)
                {
                    Marshal.FinalReleaseComObject(_outlookApp);
                    _outlookApp = null;
                }
            }
            catch (Exception ex)
            {
                ShowError("OnDisconnection failed", ex);
            }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string ribbonId)
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                using (var stream = asm.GetManifestResourceStream(
                    "RBLclass.AddIn.Ribbon.xml"))
                {
                    if (stream == null)
                        return string.Empty;
                    using (var reader = new System.IO.StreamReader(stream))
                        return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                ShowError("GetCustomUI failed", ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// Step 0 smoke-test action: proves the add-in loaded, the ribbon
        /// callback resolves, and the Core reference is wired. Replaced by the
        /// real "Open folder" / "Classify" actions in Steps 3-4.
        /// </summary>
        public void OnAboutClick(Office.IRibbonControl control)
        {
            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                var message =
                    Core.ProductInfo.Name + " add-in loaded." + Environment.NewLine +
                    "Version        : " + version + Environment.NewLine +
                    "Is64BitProcess : " + Environment.Is64BitProcess + Environment.NewLine +
                    "IntPtr.Size    : " + IntPtr.Size;

                MessageBox.Show(message, "RBLclass",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError("About action failed", ex);
            }
        }

        private static void ShowError(string title, Exception ex)
        {
            try
            {
                MessageBox.Show(ex.ToString(), "RBLclass - " + title,
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch
            {
                // Never let an exception escape into Outlook - it would
                // disable the add-in by flipping LoadBehavior 3 -> 2.
            }
        }
    }
}
