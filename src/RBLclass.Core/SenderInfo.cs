namespace RBLclass.Core
{
    /// <summary>
    /// Envelope identity of a mail's sender, for the phishing warnings (v2.13.0.0 D2):
    /// display name and real SMTP address. Either part may be empty.
    /// </summary>
    public sealed class SenderInfo
    {
        public static readonly SenderInfo Unknown = new SenderInfo(string.Empty, string.Empty);

        public SenderInfo(string displayName, string smtpAddress)
        {
            DisplayName = displayName ?? string.Empty;
            SmtpAddress = smtpAddress ?? string.Empty;
        }

        public string DisplayName { get; }
        public string SmtpAddress { get; }

        /// <summary>True when at least the name or the address is known.</summary>
        public bool IsKnown => DisplayName.Length > 0 || SmtpAddress.Length > 0;
    }
}
