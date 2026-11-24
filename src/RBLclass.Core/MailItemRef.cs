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

        /// <summary>
        /// A filed item whose new EntryID could not be identified (v2.14.0.0 A2:
        /// a protected mail moved at store level, which returns no new ID). It
        /// counts as filed, but it cannot be undone, copied or re-checked.
        /// </summary>
        public static MailItemRef Untracked(string storeId, string subject = null) =>
            new MailItemRef(storeId, string.Empty, subject);

        /// <summary>
        /// A protected mail filed at store level (v2.14.0.0 A2) whose stored row
        /// was read back from the destination's contents table as still encrypted.
        /// Re-checking it through the object model is unreliable: Outlook can hand
        /// back the item it holds in memory (decrypted, with a stale parent).
        /// </summary>
        public static MailItemRef VerifiedProtected(string storeId, string entryId, string subject = null) =>
            new MailItemRef(storeId, entryId, subject) { ProtectionVerified = true };

        /// <summary>True for <see cref="VerifiedProtected"/>: no post-filing protection check is needed.</summary>
        public bool ProtectionVerified { get; private set; }

        public string StoreId { get; }
        public string EntryId { get; }

        /// <summary>False for <see cref="Untracked"/>: there is no EntryID to resolve.</summary>
        public bool IsTracked => EntryId.Length > 0;

        /// <summary>Subject line, for display in the classify list. May be empty.</summary>
        public string Subject { get; }

        /// <summary>
        /// True when this ref was captured from a MeetingItem rather than a MailItem.
        /// The adapter handles the two COM types differently; the Core classifier passes it through unchanged.
        /// </summary>
        public bool IsMeetingItem { get; }
    }
}
