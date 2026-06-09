using System;
using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// Legacy 6a: warn before sending to recipients outside the
    /// organisation. ⚠ DEVIATION (correctness): the legacy decides
    /// internal-vs-external by an `Address` substring test for the
    /// hard-coded marker "CORPMAIL" - brittle and environment-specific.
    /// A recipient counts as internal here when Outlook resolved it against
    /// the local Exchange org (<see cref="RecipientAddress.IsExchangeResolved"/>),
    /// or when its address domain matches the configurable allowlist (empty
    /// by default - nothing org-specific is hard-coded; Step 9 exposes it).
    /// </summary>
    public sealed class ExternalRecipientGuard
    {
        public IReadOnlyList<RecipientAddress> FindExternal(IReadOnlyList<RecipientAddress> recipients,
                                                             IReadOnlyList<string> internalDomains)
        {
            if (recipients == null) throw new ArgumentNullException(nameof(recipients));

            var domains = internalDomains ?? new string[0];
            return recipients.Where(r => !IsInternal(r, domains)).ToArray();
        }

        private static bool IsInternal(RecipientAddress recipient, IReadOnlyList<string> internalDomains)
        {
            if (recipient.IsExchangeResolved) return true;

            string domain = DomainOf(recipient.Address);
            return domain != null &&
                   internalDomains.Any(d => string.Equals(d, domain, StringComparison.OrdinalIgnoreCase));
        }

        private static string DomainOf(string address)
        {
            if (string.IsNullOrEmpty(address)) return null;
            int at = address.LastIndexOf('@');
            return at >= 0 && at < address.Length - 1 ? address.Substring(at + 1) : null;
        }
    }
}
