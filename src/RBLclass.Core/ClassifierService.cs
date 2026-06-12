using System;
using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// Classify orchestration, decoupled from Outlook. For each item: with
    /// "keep a copy" on, a copy is placed in every destination and the original
    /// stays put; with it off, copies go to every destination but the last and
    /// the original is <b>moved</b> there (v2.2 - the old copy-then-delete made
    /// other add-ins, notably Stormshield, chase transient items and throw
    /// MAPI_E_NOT_FOUND; a single destination is now a pure move). Attachment
    /// stripping / task completion act on each filed item, never on a kept
    /// original. Each item is handled independently so one failure does not
    /// abort the batch.
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

            var skippedEncrypted = new List<string>();
            var widened = widenConversation ? Widen(items, skippedEncrypted) : Dedupe(items);

            var flagged = widened.Where(i => SafeIsFlaggedIncomplete(i)).ToArray();

            return new ClassifyPreflight(widened, flagged, skippedEncrypted);
        }

        public ClassifyResult Classify(ClassifyRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));

            // Nothing to do without both items and at least one destination.
            if (request.Items.Count == 0 || request.Destinations.Count == 0)
                return new ClassifyResult(0, 0, 0, 0);

            int processed = 0, copies = 0, moved = 0, errors = 0, encryptedSkips = 0;

            // Deleted Items per source store, resolved at most once per call
            // (only needed for the opt-in safety copy). Misses are cached too.
            Dictionary<string, FolderNode> deletedItemsByStore = null;

            foreach (var item in request.Items)
            {
                try
                {
                    bool markComplete = request.MarkTasksComplete && SafeIsFlaggedIncomplete(item);
                    int filed = 0;

                    // Copies first (the original must stay put while it is the
                    // copy source); the original itself is then moved into the
                    // LAST destination unless "keep a copy" is on. After the
                    // move the original reference is invalid, so the move slot
                    // must come last.
                    int count = request.Destinations.Count;
                    for (int d = 0; d < count; d++)
                    {
                        var destination = request.Destinations[d];
                        bool moveSlot = !request.KeepCopy && d == count - 1;

                        // Isolate each destination: one that refuses the item
                        // (e.g. a store root that can't hold mail) must not abort
                        // filing into the others.
                        try
                        {
                            MailItemRef filedRef;
                            if (moveSlot)
                            {
                                filedRef = _store.MoveItemToFolder(item, destination);
                                if (filedRef == null)
                                {
                                    // The item no longer resolves - the original
                                    // was not moved anywhere.
                                    errors++;
                                    continue;
                                }
                                moved++;

                                // Opt-in guardrail: leave a copy in the source
                                // store's Deleted Items, taken from the moved
                                // item at its destination (never a transient in
                                // the displayed folder) and BEFORE stripping so
                                // the guardrail copy keeps its attachments. Its
                                // failure never fails the filing itself.
                                if (request.SafetyCopy)
                                {
                                    try
                                    {
                                        var deletedItems = ResolveDeletedItems(
                                            item.StoreId, ref deletedItemsByStore);
                                        if (deletedItems != null)
                                            _store.CopyItemToFolder(filedRef, deletedItems);
                                    }
                                    catch { /* guardrail only; the adapter logs the cause */ }
                                }
                            }
                            else
                            {
                                filedRef = _store.CopyItemToFolder(item, destination);
                                copies++;
                            }
                            filed++;

                            // Strip attachments / mark the task complete on the
                            // FILED item only - with "keep a copy" on, the kept
                            // original (and its flag/attachments) stays untouched.
                            // Encrypted items are never stripped (their
                            // attachments ARE the message); the store reports
                            // the skip so the UI can say so.
                            if (filedRef != null)
                            {
                                if (request.RemoveAttachments && !_store.RemoveAttachments(filedRef))
                                    encryptedSkips++;
                                if (markComplete)
                                    _store.MarkTaskComplete(filedRef);
                            }
                        }
                        catch
                        {
                            // This destination failed; the adapter logs the cause.
                            // A failed move leaves the original in place - never
                            // any data loss on failure.
                            errors++;
                        }
                    }

                    if (filed > 0)
                        processed++;
                }
                catch
                {
                    // Unexpected per-item failure (e.g. the flag check); keep going.
                    errors++;
                }
            }

            return new ClassifyResult(processed, copies, moved, errors, encryptedSkips);
        }

        /// <summary>
        /// The original items plus every conversation sibling reported by the
        /// store (legacy 5b step 2), deduped by (StoreId, EntryId) - the
        /// originals win the slot when a sibling lookup also returns one of
        /// them back.
        /// </summary>
        private IReadOnlyList<MailItemRef> Widen(IReadOnlyList<MailItemRef> items,
                                                 List<string> skippedEncrypted)
        {
            var seen = new HashSet<(string StoreId, string EntryId)>();
            var result = new List<MailItemRef>();

            void AddIfNew(MailItemRef candidate)
            {
                if (candidate != null && seen.Add((candidate.StoreId, candidate.EntryId)))
                    result.Add(candidate);
            }

            // Dedupe skipped-encrypted reports too: the same encrypted sibling can
            // be reported by more than one source item in the selection.
            var skippedSeen = new HashSet<string>();

            foreach (var item in items) AddIfNew(item);
            foreach (var item in items)
            {
                var siblings = _store.GetConversationSiblings(item) ?? ConversationSiblings.Empty;
                foreach (var sibling in siblings.Processable) AddIfNew(sibling);
                foreach (var subject in siblings.SkippedEncryptedSubjects)
                    if (skippedSeen.Add(subject)) skippedEncrypted.Add(subject);
            }

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

        private FolderNode ResolveDeletedItems(string storeId,
                                               ref Dictionary<string, FolderNode> cache)
        {
            if (cache == null) cache = new Dictionary<string, FolderNode>();
            FolderNode node;
            if (!cache.TryGetValue(storeId, out node))
            {
                try { node = _store.GetDeletedItemsFolder(storeId); }
                catch { node = null; }
                cache[storeId] = node; // cache misses too
            }
            return node;
        }

        private bool SafeIsFlaggedIncomplete(MailItemRef item)
        {
            try { return _store.IsFlaggedIncomplete(item); }
            catch { return false; } // tolerate adapter misses; never block the batch
        }
    }
}
