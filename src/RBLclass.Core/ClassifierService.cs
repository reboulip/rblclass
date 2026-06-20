using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

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
        private readonly IClassificationHistory _history;

        /// <param name="history">
        /// Optional: when provided, every successful filing is recorded against
        /// the item's conversation (feeding Auto-class) and Undo rolls those
        /// records back. Recording is best-effort - it never fails a classify.
        /// </param>
        public ClassifierService(IMailStore store, IClassificationHistory history = null)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));
            _history = history;
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

            var acc = new ClassifyAccumulator();
            foreach (var item in request.Items)
                ProcessItem(item, request, acc);

            return BuildResult(acc, WriteHistory(request, acc));
        }

        public async Task<ClassifyResult> ClassifyAsync(ClassifyRequest request,
            IProgress<ClassifyProgress> progress, Func<Task> yieldBetweenItems)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));

            if (request.Items.Count == 0 || request.Destinations.Count == 0)
                return new ClassifyResult(0, 0, 0, 0);

            var acc = new ClassifyAccumulator();
            int total = request.Items.Count;
            for (int i = 0; i < total; i++)
            {
                ProcessItem(request.Items[i], request, acc);
                progress?.Report(new ClassifyProgress(i + 1, total));

                // Yield the caller's message pump between items (not after the
                // last - the history await below already releases it). Repaints
                // Outlook, keeps it interactive, and gives a mail scanner a
                // stable window before the next move (D1).
                if (yieldBetweenItems != null && i < total - 1)
                    await yieldBetweenItems();
            }

            // The history write touches SQLite, not COM - run it off the
            // caller's (UI) thread.
            var historyBatchIds = await Task.Run(() => WriteHistory(request, acc));
            return BuildResult(acc, historyBatchIds);
        }

        /// <summary>
        /// File one item per <paramref name="request"/>, accumulating reversible
        /// actions and counts into <paramref name="acc"/>. Shared by the
        /// synchronous <see cref="Classify"/> and the responsive
        /// <see cref="ClassifyAsync"/> so both produce an identical item order
        /// and undo plan.
        /// </summary>
        private void ProcessItem(MailItemRef item, ClassifyRequest request, ClassifyAccumulator acc)
        {
            try
            {
                bool markComplete = request.MarkTasksComplete && SafeIsFlaggedIncomplete(item);
                int filed = 0;

                // Where the item lives now - so Undo can put a moved
                // original back. Resolved before anything moves it.
                FolderNode sourceFolder = null;
                if (!request.KeepCopy)
                {
                    try { sourceFolder = _store.GetParentFolder(item); }
                    catch { /* tolerated: the move just won't be undoable */ }
                }

                // Conversation key for the history - read before the move
                // invalidates the reference (only when history is enabled).
                string conversationKey = null;
                if (_history != null)
                {
                    try { conversationKey = _store.GetConversationKey(item); }
                    catch { /* tolerated: this item just won't feed Auto-class */ }
                }

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
                                acc.Errors++;
                                continue;
                            }
                            acc.Moved++;
                            if (sourceFolder != null)
                                acc.Moves.Add(new UndoableMove(filedRef, sourceFolder));

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
                                        item.StoreId, ref acc.DeletedItemsByStore);
                                    if (deletedItems != null)
                                    {
                                        var safety = _store.CopyItemToFolder(filedRef, deletedItems);
                                        if (safety != null) acc.CreatedCopies.Add(safety);
                                    }
                                }
                                catch { /* guardrail only; the adapter logs the cause */ }
                            }
                        }
                        else
                        {
                            filedRef = _store.CopyItemToFolder(item, destination);
                            acc.Copies++;
                            if (filedRef != null) acc.CreatedCopies.Add(filedRef);
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
                            if (request.RemoveAttachments)
                            {
                                if (_store.RemoveAttachments(filedRef)) acc.Strips++;
                                else acc.EncryptedSkips++;
                            }
                            // Strip the learned external banner from the filed
                            // body (best-effort, not undoable - a body edit).
                            if (request.StripBanner)
                            {
                                try { _store.StripExternalBanner(filedRef, request.BannerSignature); }
                                catch { /* adapter logs; never fail the filing over a banner */ }
                            }
                            if (markComplete)
                            {
                                _store.MarkTaskComplete(filedRef);
                                acc.Flags.Add(filedRef);
                            }
                        }
                    }
                    catch
                    {
                        // This destination failed; the adapter logs the cause.
                        // A failed move leaves the original in place - never
                        // any data loss on failure.
                        acc.Errors++;
                    }
                }

                if (filed > 0)
                {
                    acc.Processed++;
                    if (conversationKey != null && conversationKey.Length > 0)
                        acc.FiledConversationKeys.Add(conversationKey);
                }
            }
            catch
            {
                // Unexpected per-item failure (e.g. the flag check); keep going.
                acc.Errors++;
            }
        }

        /// <summary>
        /// Record the history for Auto-class: one batch for the whole classify
        /// action, one row per (filed conversation, destination). Best-effort -
        /// a history write never fails the classify. Returns the batch id(s).
        /// </summary>
        private List<string> WriteHistory(ClassifyRequest request, ClassifyAccumulator acc)
        {
            var historyBatchIds = new List<string>();
            if (_history != null && acc.FiledConversationKeys.Count > 0)
            {
                string batchId = Guid.NewGuid().ToString("N");
                try
                {
                    var whenUtc = DateTime.UtcNow;
                    foreach (var key in acc.FiledConversationKeys)
                        _history.Record(batchId, key, request.Destinations, whenUtc);
                    historyBatchIds.Add(batchId);
                }
                catch { /* Auto-class just won't learn this one */ }
            }
            return historyBatchIds;
        }

        private static ClassifyResult BuildResult(ClassifyAccumulator acc, List<string> historyBatchIds)
        {
            var plan = new ClassifyUndoPlan(acc.Moves, acc.CreatedCopies, acc.Flags, acc.Strips, historyBatchIds);
            return new ClassifyResult(acc.Processed, acc.Copies, acc.Moved, acc.Errors, acc.EncryptedSkips,
                                      plan.IsEmpty ? null : plan);
        }

        /// <summary>Mutable per-call scratch shared by the sync and async classify paths.</summary>
        private sealed class ClassifyAccumulator
        {
            public int Processed, Copies, Moved, Errors, EncryptedSkips, Strips;

            // Everything reversible, recorded as it happens (v2.2 Undo).
            public readonly List<UndoableMove> Moves = new List<UndoableMove>();
            public readonly List<MailItemRef> CreatedCopies = new List<MailItemRef>();
            public readonly List<MailItemRef> Flags = new List<MailItemRef>();

            // Conversation keys of items actually filed, for the Auto-class
            // history (deduped - several selected items can share a thread).
            public readonly HashSet<string> FiledConversationKeys = new HashSet<string>();

            // Deleted Items per source store, resolved at most once per call
            // (only needed for the opt-in safety copy). Misses are cached too.
            public Dictionary<string, FolderNode> DeletedItemsByStore;
        }

        public UndoResult Undo(ClassifyUndoPlan plan)
        {
            if (plan == null) throw new ArgumentNullException(nameof(plan));

            int movesRestored = 0, copiesDeleted = 0, flagsRestored = 0, errors = 0;

            // Flags first: their references point at the filed items where they
            // are NOW - they become stale once the moves below put items back.
            foreach (var flagged in plan.CompletedFlags)
            {
                try { _store.MarkTaskIncomplete(flagged); flagsRestored++; }
                catch { errors++; }
            }

            // Copies the classify created are removed outright ("undo
            // completely" - they were never user data, just our duplicates).
            foreach (var copy in plan.CreatedCopies)
            {
                try { _store.DeleteItemPermanently(copy); copiesDeleted++; }
                catch { errors++; }
            }

            // Finally put each moved original back where it came from.
            foreach (var move in plan.Moves)
            {
                try
                {
                    if (_store.MoveItemToFolder(move.Current, move.SourceFolder) != null)
                        movesRestored++;
                    else
                        errors++;
                }
                catch { errors++; }
            }

            // Roll back what Auto-class learned from this classify (best-effort:
            // stale history would only mis-suggest, never lose data).
            if (_history != null && plan.HistoryBatchIds.Count > 0)
            {
                try { _history.DeleteBatches(plan.HistoryBatchIds); }
                catch { /* leave the history rows; they can be re-classified over */ }
            }

            return new UndoResult(movesRestored, copiesDeleted, flagsRestored, errors);
        }

        public AutoClassifyResult AutoClassify(IReadOnlyList<MailItemRef> items,
                                               Func<string, string, FolderNode> resolveLiveFolder,
                                               bool keepCopy, bool removeAttachments, bool safetyCopy)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (resolveLiveFolder == null) throw new ArgumentNullException(nameof(resolveLiveFolder));
            if (_history == null)
                throw new InvalidOperationException("Auto-class requires a classification history store.");

            int filed = 0, noHistory = 0, staleFolders = 0, errors = 0;

            // One combined undo plan over every per-item filing this run does.
            var allMoves = new List<UndoableMove>();
            var allCopies = new List<MailItemRef>();
            var allFlags = new List<MailItemRef>();
            int strips = 0;
            var allBatchIds = new List<string>();

            // Distinct live folders we actually filed into, to show the user.
            var filedDestinations = new List<FolderNode>();
            var filedDestKeys = new HashSet<string>();

            foreach (var item in Dedupe(items))
            {
                string key;
                try { key = _store.GetConversationKey(item); }
                catch { key = null; }

                if (string.IsNullOrEmpty(key)) { noHistory++; continue; }

                var recorded = _history.GetLatestDestinations(key);
                if (recorded.Count == 0) { noHistory++; continue; }

                // Validate each remembered destination against the live index;
                // a folder deleted or renamed away no longer resolves.
                var live = new List<FolderNode>();
                foreach (var dest in recorded)
                {
                    FolderNode node;
                    try { node = resolveLiveFolder(dest.StoreId, dest.EntryId); }
                    catch { node = null; }
                    if (node != null) live.Add(node);
                }

                if (live.Count == 0) { staleFolders++; continue; }

                // Reuse the full classify path - same move/copy/strip/history
                // semantics, and it records a fresh batch that Undo rolls back.
                var result = Classify(new ClassifyRequest(
                    new[] { item }, live, keepCopy, removeAttachments,
                    markTasksComplete: false, safetyCopy: safetyCopy));

                if (result.ItemsProcessed > 0)
                {
                    filed++;
                    foreach (var dest in live)
                        if (filedDestKeys.Add(dest.StoreId + " " + dest.EntryId))
                            filedDestinations.Add(dest);
                    if (result.Undo != null)
                    {
                        allMoves.AddRange(result.Undo.Moves);
                        allCopies.AddRange(result.Undo.CreatedCopies);
                        allFlags.AddRange(result.Undo.CompletedFlags);
                        strips += result.Undo.AttachmentStrips;
                        allBatchIds.AddRange(result.Undo.HistoryBatchIds);
                    }
                }
                else
                {
                    errors++;
                }
            }

            var plan = new ClassifyUndoPlan(allMoves, allCopies, allFlags, strips, allBatchIds);
            return new AutoClassifyResult(filed, noHistory, staleFolders, errors,
                                          plan.IsEmpty ? null : plan, filedDestinations);
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
