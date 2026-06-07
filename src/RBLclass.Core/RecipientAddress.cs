namespace RBLclass.Core
{
    /// <summary>
    /// One recipient of an outgoing mail, as seen by the send-time guards
    /// (legacy 6a). <see cref="IsExchangeResolved"/> carries the only signal
    /// the adapter can extract without a configurable domain list: whether
    /// Outlook resolved the address against the local Exchange org.
    /// </summary>
    public sealed class RecipientAddress
    {
        public RecipientAddress(string displayName, string address, bool isExchangeResolved)
        {
            DisplayName = displayName;
            Address = address;
            IsExchangeResolved = isExchangeResolved;
        }

        public string DisplayName { get; }

        public string Address { get; }

        /// <summary>
        /// True when Outlook resolved this address as belonging to the local
        /// Exchange org (a user, distribution list, public folder, ...) -
        /// the adapter's <c>AddressEntryUserType</c> reading. Treated as the
        /// strong "internal" signal; everything else falls back to the
        /// (currently empty-by-default) internal-domain allowlist.
        /// </summary>
        public bool IsExchangeResolved { get; }
    }
}
