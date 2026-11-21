using RBLclass.AddIn.Localization;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// Text of the "Possible phishing" warning (v2.13.0.0 D2). Shared by the attachment
    /// warning and, later, the link warning: every line is optional and its block is
    /// hidden when empty.
    /// </summary>
    public sealed class PhishingWarningViewModel
    {
        public string Text { get; set; }

        /// <summary>"From: name &lt;address&gt;", or just what is known; empty when nothing is.</summary>
        public string SenderLine { get; set; }

        /// <summary>The attachment file name line (attachment mode), empty when unknown.</summary>
        public string DetailLine { get; set; }

        /// <summary>The full link destination (link mode), empty otherwise.</summary>
        public string UrlText { get; set; }

        /// <summary>"Shown in the email as: …" when the displayed text differs from the URL.</summary>
        public string DisplayedTextLine { get; set; }

        public bool HasSender => !string.IsNullOrEmpty(SenderLine);
        public bool HasDetail => !string.IsNullOrEmpty(DetailLine);
        public bool HasUrl => !string.IsNullOrEmpty(UrlText);
        public bool HasDisplayedText => !string.IsNullOrEmpty(DisplayedTextLine);

        /// <summary>The sender line from the envelope identity, or empty when nothing is known.</summary>
        public static string FormatSender(ILocalizationService loc, SenderInfo sender)
        {
            if (sender == null || !sender.IsKnown) return string.Empty;
            if (sender.DisplayName.Length > 0 && sender.SmtpAddress.Length > 0)
                return loc.GetString("PhishingLink_Sender", sender.DisplayName, sender.SmtpAddress);
            return sender.DisplayName.Length > 0 ? sender.DisplayName : sender.SmtpAddress;
        }

        public static PhishingWarningViewModel ForLink(ILocalizationService loc, LinkWarningMessage message)
        {
            return new PhishingWarningViewModel
            {
                Text = loc.GetString("PhishingLink_Text"),
                SenderLine = FormatSender(loc, message.Sender),
                DetailLine = string.Empty,
                UrlText = loc.GetString("PhishingLink_Url", message.Url),
                DisplayedTextLine = message.DisplayedTextDiffers
                    ? loc.GetString("PhishingLink_DisplayedText", message.TextToDisplay)
                    : string.Empty
            };
        }

        public static PhishingWarningViewModel ForAttachment(ILocalizationService loc, PhishingWarningMessage message)
        {
            return new PhishingWarningViewModel
            {
                Text = loc.GetString("PhishingAttachment_Text"),
                SenderLine = FormatSender(loc, message.Sender),
                DetailLine = message.HasAttachmentFileName
                    ? loc.GetString("PhishingAttachment_FileName", message.AttachmentFileName)
                    : string.Empty,
                UrlText = string.Empty,
                DisplayedTextLine = string.Empty
            };
        }
    }
}
