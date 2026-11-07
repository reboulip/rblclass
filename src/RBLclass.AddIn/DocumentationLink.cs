using System;
using System.Diagnostics;
using Serilog;

namespace RBLclass.AddIn
{
    /// <summary>
    /// Opens the online documentation in the default browser (v2.11.0.0 B2). The
    /// add-in itself sends nothing: the browser does the navigation.
    /// </summary>
    internal static class DocumentationLink
    {
        private const string Url = "https://reboulip.github.io/rblclass/";

        /// <summary>Returns false, after logging, when no browser could be started.</summary>
        internal static bool TryOpen()
        {
            try
            {
                using (Process.Start(new ProcessStartInfo(Url) { UseShellExecute = true })) { }
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Opening the online documentation failed.");
                return false;
            }
        }
    }
}
