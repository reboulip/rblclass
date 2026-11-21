using System;

namespace RBLclass.Core
{
    /// <summary>
    /// Which link targets deserve the "possible phishing" warning, and how a link is
    /// written to the log (v2.13.0.0 D3). Pure and Outlook-free.
    /// </summary>
    public static class PhishingLinkFilter
    {
        /// <summary>
        /// False for a blank target, an in-document anchor (<c>#section</c>) and
        /// <c>mailto:</c> (it opens a compose window, not a page). True for everything
        /// else: http, https, ftp, file and unknown schemes.
        /// </summary>
        public static bool IsWarnable(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return false;
            string trimmed = url.Trim();
            if (trimmed.StartsWith("#", StringComparison.Ordinal)) return false;
            if (trimmed.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) return false;
            return true;
        }

        /// <summary>
        /// <c>scheme://host</c> for the log, never the path or query (they can carry
        /// tokens); <c>(unparsed)</c> when the target is not an absolute URL.
        /// </summary>
        public static string HostForLog(string url)
        {
            Uri uri;
            if (string.IsNullOrWhiteSpace(url) || !Uri.TryCreate(url.Trim(), UriKind.Absolute, out uri))
                return "(unparsed)";

            string host = uri.Host;
            if (string.IsNullOrEmpty(host)) return uri.Scheme + "://";
            return uri.Scheme + "://" + host;
        }
    }
}
