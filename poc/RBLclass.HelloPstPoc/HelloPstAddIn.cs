using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RBLclass.HelloPstPoc
{
    [ComVisible(true)]
    [Guid(ClsidString)]
    [ProgId(ProgIdString)]
    [ClassInterface(ClassInterfaceType.None)]
    public class HelloPstAddIn : IDTExtensibility2, Office.IRibbonExtensibility
    {
        public const string ClsidString = "B5E7A6F1-3D5A-4A8B-9C2E-1F4E5A8B9D03";
        public const string ProgIdString = "RBLclass.HelloPstAddIn";

        private Outlook.Application _outlookApp;

        public void OnConnection(object application,
                                 ext_ConnectMode connectMode,
                                 object addInInst,
                                 ref Array custom)
        {
            try
            {
                _outlookApp = (Outlook.Application)application;
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
                    "RBLclass.HelloPstPoc.Ribbon.xml"))
                {
                    if (stream == null)
                        return string.Empty;
                    using (var reader = new StreamReader(stream))
                        return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                ShowError("GetCustomUI failed", ex);
                return string.Empty;
            }
        }

        public void OnHelloPstClick(Office.IRibbonControl control)
        {
            try
            {
                HelloPstAction.Run(_outlookApp);
            }
            catch (Exception ex)
            {
                ShowError("Hello PST action failed", ex);
            }
        }

        private static void ShowError(string title, Exception ex)
        {
            try
            {
                MessageBox.Show(ex.ToString(), "RBLclass POC — " + title,
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch
            {
                // Never let an exception escape into Outlook — it would
                // disable the add-in by flipping LoadBehavior 3 -> 2.
            }
        }
    }
}
