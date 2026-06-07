using System;

namespace RBLclass.Core
{
    /// <summary>
    /// A stable reference to a mail item by (StoreId, EntryId), with an optional
    /// subject for display. The classifier works purely with these refs; the
    /// adapter resolves them to live Outlook items.
    /// </summary>
    public sealed class MailItemRef
    {
        public MailItemRef(string storeId, string entryId, string subject = null)
        {
            StoreId = storeId ?? throw new ArgumentNullException(nameof(storeId));
            EntryId = entryId ?? throw new ArgumentNullException(nameof(entryId));
            Subject = subject ?? string.Empty;
        }

        public string StoreId { get; }
        public string EntryId { get; }

        /// <summary>Subject line, for display in the classify list. May be empty.</summary>
        public string Subject { get; }
    }
}
