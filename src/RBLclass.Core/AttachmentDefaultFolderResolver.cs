using System;
using System.Threading.Tasks;

namespace RBLclass.Core
{
    /// <summary>
    /// Decides whether the configured default attachment folder can pre-fill the
    /// disposition modal (v2.11.0.0 C1). An unreachable folder, such as a
    /// disconnected share, leaves the field empty. The existence check can block
    /// for a long time on a dead network path, so it is bounded by a timeout.
    /// </summary>
    public static class AttachmentDefaultFolderResolver
    {
        /// <summary>
        /// The trimmed folder when <paramref name="exists"/> confirms it within
        /// <paramref name="timeout"/>, otherwise an empty string. Blocks for up to
        /// the timeout: call it off the UI thread.
        /// </summary>
        public static string Resolve(string configured, Func<string, bool> exists, TimeSpan timeout)
        {
            if (string.IsNullOrWhiteSpace(configured) || exists == null) return string.Empty;

            string path = configured.Trim();
            try
            {
                var probe = Task.Run(() => exists(path));
                return probe.Wait(timeout) && probe.Result ? path : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
