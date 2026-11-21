using System;

namespace RBLclass.Core
{
    /// <summary>
    /// What the link warning shows (v2.13.0.0 D3): the sender, the full destination and
    /// the text the mail displays for the link. A displayed text that differs from the
    /// destination is a strong phishing signal and gets its own line.
    /// </summary>
    public sealed class LinkWarningMessage
    {
        public LinkWarningMessage(SenderInfo sender, string url, string textToDisplay)
        {
            Sender = sender ?? SenderInfo.Unknown;
            Url = url ?? string.Empty;
            TextToDisplay = textToDisplay ?? string.Empty;
        }

        public SenderInfo Sender { get; }
        public string Url { get; }
        public string TextToDisplay { get; }

        /// <summary>
        /// True when the mail shows something other than the destination for this link
        /// (ignoring case, surrounding spaces and a trailing slash).
        /// </summary>
        public bool DisplayedTextDiffers
        {
            get
            {
                string shown = TextToDisplay.Trim();
                if (shown.Length == 0) return false;
                return !string.Equals(Normalize(shown), Normalize(Url), StringComparison.OrdinalIgnoreCase);
            }
        }

        private static string Normalize(string s) => (s ?? string.Empty).Trim().TrimEnd('/');
    }
}
