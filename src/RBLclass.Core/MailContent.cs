using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// The decrypted content of a protected mail, as seen in memory, that
    /// re-encrypted filing rebuilds (v2.12.0.0 A2). The envelope (subject,
    /// sender, recipients, dates) is not part of it: it stays outside the S/MIME
    /// encryption and is copied by the adapter. Outlook-free.
    /// </summary>
    public sealed class MailContent
    {
        public MailContent(string htmlBody, string plainBody, IReadOnlyList<MailPart> parts)
        {
            HtmlBody = htmlBody;
            PlainBody = plainBody;
            Parts = parts == null ? new MailPart[0] : parts.Where(p => p != null).ToArray();
        }

        /// <summary>The HTML body, or null when the mail has none.</summary>
        public string HtmlBody { get; }

        /// <summary>The plain-text body, or null when the mail has none.</summary>
        public string PlainBody { get; }

        public IReadOnlyList<MailPart> Parts { get; }

        /// <summary>Sum of every part's byte count (size cap, logging).</summary>
        public long TotalPartBytes
        {
            get
            {
                long total = 0;
                foreach (var part in Parts) total += part.Data.Length;
                return total;
            }
        }
    }
}
