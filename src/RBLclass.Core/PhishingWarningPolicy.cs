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
    }
}
