using System;

namespace RBLclass.Core
{
    /// <summary>
    /// Decides whether opening an item from a mail carrying the organization's
    /// external-sender banner needs the "possible phishing" warning (v2.13.0.0 D2).
    /// Pure and Outlook-free. The body is only ever read when the event fires.
    /// </summary>
    public static class PhishingWarningPolicy
    {
        /// <summary>
        /// Warn before an attachment opens when the setting is on, a banner signature
        /// has been learned, and the mail's HTML body carries that banner. No warning
        /// when no signature has been learned.
        /// </summary>
        public static bool ShouldWarnAttachment(bool settingOn, string bannerSignature, string htmlBody)
        {
            if (!settingOn) return false;
            if (string.IsNullOrWhiteSpace(bannerSignature)) return false;
            if (string.IsNullOrEmpty(htmlBody)) return false;
            return ExternalBannerStripper.ContainsBanner(htmlBody, bannerSignature);
        }

        /// <summary>
        /// Warn before a link opens when the setting is on, a banner signature has been
        /// learned, the target is one worth warning about (not an anchor or mailto:) and
        /// the mail's RENDERED text carries the banner. The rendered text is used (not the
        /// HTML) because the stored HTML may be in a character set that breaks the match.
        /// (v2.13.0.0 D3)
        /// </summary>
        public static bool ShouldWarnLink(bool settingOn, string bannerSignature, string renderedText, string url)
        {
            if (!settingOn) return false;
            if (string.IsNullOrWhiteSpace(bannerSignature)) return false;
            if (!PhishingLinkFilter.IsWarnable(url)) return false;
            return ExternalBannerStripper.ContainsBannerText(renderedText, bannerSignature);
        }
    }
}
