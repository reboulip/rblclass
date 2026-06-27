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
        public MailItemRef(string storeId, string entryId,
                           string subject = null, bool isMeetingItem = false)
        {
            StoreId = storeId ?? throw new ArgumentNullException(nameof(storeId));
            EntryId = entryId ?? throw new ArgumentNullException(nameof(entryId));
            Subject = subject ?? string.Empty;
            IsMeetingItem = isMeetingItem;
        }

        public string StoreId { get; }
        public string EntryId { get; }

        /// <summary>Subject line, for display in the classify list. May be empty.</summary>
        public string Subject { get; }

        /// <summary>
        /// True when this ref was captured from a MeetingItem rather than a MailItem.
        /// The adapter handles the two COM types differently; the Core classifier passes it through unchanged.
        /// </summary>
        public bool IsMeetingItem { get; }
    }
}
