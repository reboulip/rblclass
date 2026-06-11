using System;
using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// Classify orchestration, decoupled from Outlook. For each item: optionally
    /// strip attachments, copy it into every destination (legacy
    /// copy-per-destination), then delete the original unless "keep a copy" is
    /// on. Each item is handled independently so one failure does not abort the
    /// batch.
    /// </summary>
    public sealed class ClassifierService : IClassifier
    {
        private readonly IMailStore _store;

        public ClassifierService(IMailStore store)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));
        }

        public ClassifyPreflight Preflight(IReadOnlyList<MailItemRef> items, bool widenConversation)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));

            var widened = widenConversation ? Widen(items) : Dedupe(items);

            var flagged = widened.Where(i => SafeIsFlaggedIncomplete(i)).ToArray();

            return new ClassifyPreflight(widened, flagged);
        }

        public ClassifyResult Classify(ClassifyRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));

            // Nothing to do without both items and at least one destination.
            if (request.Items.Count == 0 || request.Destinations.Count == 0)
                return new ClassifyResult(0, 0, 0, 0);

            int processed = 0, copies = 0, deleted = 0, errors = 0;

            foreach (var item in request.Items)
            {
                try
                {
                    bool markComplete = request.MarkTasksComplete && SafeIsFlaggedIncomplete(item);
                    int filed = 0;

                    foreach (var destination in request.Destinations)
                    {
                        // Isolate each destination: one that refuses the item
                        // (e.g. a store root that can't hold mail) must not abort
                        // filing into the others.
                        try
                        {
                            var copy = _store.CopyItemToFolder(item, destination);
                            copies++;
                            filed++;

                            // Strip attachments / mark the task complete on the
                            // FILED COPY only, never the original - so "keep a
                            // copy" leaves the original (and its flag/attachments)
                            // untouched.
                            if (copy != null)
                            {
                                if (request.RemoveAttachments)
                                    _store.RemoveAttachments(copy);
                                if (markComplete)
                                    _store.MarkTaskComplete(copy);
                            }
                        }
                        catch
                        {
                            // This destination failed; the adapter logs the cause.
                            errors++;
                        }
                    }

                    // Only delete the original once it has actually been filed
                    // somewhere; never delete a mail we could not copy to any
                    // destination (no data loss on a total failure).
                    if (filed > 0)
                    {
                        if (!request.KeepCopy)
                        {
                            _store.DeleteItem(item);
                            deleted++;
                        }
                        processed++;
                    }
                }
                catch
                {
                    // Unexpected per-item failure (e.g. the flag check); keep going.
                    errors++;
                }
            }

            return new ClassifyResult(processed, copies, deleted, errors);
        }

        /// <summary>
        /// The original items plus every conversation sibling reported by the
        /// store (legacy 5b step 2), deduped by (StoreId, EntryId) - the
        /// originals win the slot when a sibling lookup also returns one of
        /// them back.
        /// </summary>
        private IReadOnlyList<MailItemRef> Widen(IReadOnlyList<MailItemRef> items)
        {
            var seen = new HashSet<(string StoreId, string EntryId)>();
            var result = new List<MailItemRef>();

            void AddIfNew(MailItemRef candidate)
            {
                if (candidate != null && seen.Add((candidate.StoreId, candidate.EntryId)))
                    result.Add(candidate);
            }

            foreach (var item in items) AddIfNew(item);
            foreach (var item in items)
                foreach (var sibling in _store.GetConversationSiblings(item) ?? Array.Empty<MailItemRef>())
                    AddIfNew(sibling);

            return result;
        }

        private static IReadOnlyList<MailItemRef> Dedupe(IReadOnlyList<MailItemRef> items)
        {
            var seen = new HashSet<(string StoreId, string EntryId)>();
            var result = new List<MailItemRef>();
            foreach (var item in items)
                if (item != null && seen.Add((item.StoreId, item.EntryId)))
                    result.Add(item);
            return result;
        }

        private bool SafeIsFlaggedIncomplete(MailItemRef item)
        {
            try { return _store.IsFlaggedIncomplete(item); }
            catch { return false; } // tolerate adapter misses; never block the batch
        }
    }
}
