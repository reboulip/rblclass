using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using RBLclass.Core.Logging;

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
        private readonly IAppLog _log;
        private readonly IMailRebuilder _rebuilder;

        // "No usable certificate" is reported once per classifier, i.e. once per
        // Outlook session, so the status line does not repeat it on every classify.
        private bool _noCertificateReported;

        /// <param name="history">
        /// Optional: when provided, every successful filing is recorded against
        /// the item's conversation (feeding Auto-class) and Undo rolls those
        /// records back. Recording is best-effort - it never fails a classify.
        /// </param>
        /// <param name="log">Optional: defaults to a no-op sink when not provided.</param>
        /// <param name="rebuilder">
        /// Optional: enables the experimental re-encrypted filing of protected mail
        /// (v2.12.0.0) when a request asks for it. Null = the standard path only.
        /// </param>
        public ClassifierService(IMailStore store, IClassificationHistory history = null, IAppLog log = null,
                                 IMailRebuilder rebuilder = null)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));
            _history = history;
            _log = log ?? NullAppLog.Instance;
            _rebuilder = rebuilder;
        }

        public ClassifyPreflight Preflight(IReadOnlyList<MailItemRef> items, bool widenConversation)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));

            var skippedEncrypted = new List<string>();
            bool wideningTableDisabled = false;
            var widened = widenConversation
                ? Widen(items, skippedEncrypted, out wideningTableDisabled)
                : Dedupe(items);

            // Protected mail stays flagged and is not offered: completing its
            // flag would need a Save(), which it must never get (v2.11.0.0 A3).
            var flagged = widened.Where(i => SafeIsFlaggedIncomplete(i) && !SafeIsEncrypted(i)).ToArray();

            return new ClassifyPreflight(widened, flagged, skippedEncrypted, wideningTableDisabled);
        }

        public ClassifyResult Classify(ClassifyRequest request) => Classify(request, new ReencryptSession());

        /// <summary>
        /// <see cref="Classify(ClassifyRequest)"/> sharing one re-encryption session
        /// (the certificate is resolved once) across the per-item runs of Auto-class.
        /// </summary>
        private ClassifyResult Classify(ClassifyRequest request, ReencryptSession session)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));

            // Nothing to do without both items and at least one destination.
            if (request.Items.Count == 0 || request.Destinations.Count == 0)
                return new ClassifyResult(0, 0, 0, 0);

            var acc = new ClassifyAccumulator(session);
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

            var acc = new ClassifyAccumulator(new ReencryptSession());
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
                // Decided once, on the source, BEFORE anything is filed (#62).
                // Stormshield decrypts the opened mail in memory, and Outlook's
                // Move/Copy write that in-memory state into the filed item, so
                // the filed copy can look like plain mail afterwards. Its verdict
                // must not be re-derived from it: a mail protected now gets no
                // write at all on any filed copy.
                bool wasProtected = SafeIsEncrypted(item);
                bool markComplete = request.MarkTasksComplete
                    && !wasProtected && SafeIsFlaggedIncomplete(item);
                bool markAsRead = request.MarkAsRead && !wasProtected;
                int filed = 0;
                var protectionBefore = TraceProtection("before", "pending", item);

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

                // Experimental (v2.12.0.0): an encrypted mail readable decrypted in
                // memory is rebuilt, encrypted again and created as S/MIME in each
                // destination. Any failure falls through to the standard path below.
                if (wasProtected && request.ReencryptProtected && _rebuilder != null
                    && TryReencrypt(item, request, acc, sourceFolder, conversationKey))
                    return;

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
                    _log.Debug(moveSlot ? "Classify: moving item -> {Path}" : "Classify: copying item -> {Path}",
                        destination.FullPath);

                    // Isolate each destination: one that refuses the item
                    // (e.g. a store root that can't hold mail) must not abort
                    // filing into the others.
                    try
                    {
                        MailItemRef filedRef;
                        if (moveSlot)
                        {
                            filedRef = _store.MoveItemToFolder(item, destination, markAsRead);
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
                                        if (safety != null)
                                        {
                                            acc.CreatedCopies.Add(safety);
                                            TraceProtectionAfter(protectionBefore, "safety-copy", safety);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // guardrail only; the filing itself already succeeded
                                    _log.Warning(ex, "Classify safety-copy to Deleted Items failed for {Path}", destination.FullPath);
                                }
                            }
                        }
                        else
                        {
                            filedRef = _store.CopyItemToFolder(item, destination, markAsRead);
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
                        if (filedRef != null && wasProtected)
                        {
                            // No write on a protected mail's filed copy (#62).
                            // Reported like the adapter's own strip guard.
                            if (request.RemoveAttachments)
                                acc.EncryptedSkips++;
                            WarnIfProtectionLost(filedRef, moveSlot ? "move" : "copy");
                            TraceProtectionAfter(protectionBefore, moveSlot ? "move" : "copy", filedRef);
                        }
                        else if (filedRef != null)
                        {
                            // Body images chosen in the modal are handled even
                            // when attachment removal is off (v2.11.0.0 C2).
                            if (request.RemoveAttachments || request.AttachmentDispositions.Any(row => row.IsInlineImage))
                                ApplyAttachmentRemoval(item, filedRef, request, acc);
                            // Strip the learned external banner from the filed
                            // body (best-effort, not undoable - a body edit).
                            if (request.StripBanner)
                            {
                                try { _store.StripExternalBanner(filedRef, request.BannerSignature); }
                                catch { /* adapter logs; never fail the filing over a banner */ }
                            }
                            if (markComplete)
                            {
                                try
                                {
                                    _store.MarkTaskComplete(filedRef);
                                    // Record for Undo only when the flag was set.
                                    acc.Flags.Add(filedRef);
                                }
                                catch { /* adapter logs; a flag failure never fails a successful filing */ }
                            }
                            TraceProtectionAfter(protectionBefore, moveSlot ? "move" : "copy", filedRef);
                        }
                    }
                    catch (Exception ex)
                    {
                        // A failed move leaves the original in place - never
                        // any data loss on failure.
                        _log.Warning(ex, "Classify item failed -> {Path}", destination.FullPath);
                        acc.Errors++;
                    }
                }

                if (request.KeepCopy)
                    TraceProtectionAfter(protectionBefore, "kept-original", item);

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
        /// Experimental re-encrypted filing of one protected item (v2.12.0.0).
        /// Rebuilds the mail as decrypted in memory, applies the requested
        /// attachment, body-image, label and banner changes to the rebuilt content,
        /// encrypts it again and creates it directly as S/MIME in every
        /// destination, so no plaintext is stored. Unless "keep a copy" is on, the
        /// original then goes to its store's Deleted Items (deleted for good if the
        /// move wrote it decrypted). Returns false, having created nothing that
        /// stays, when the item must take the standard path instead; true when it
        /// was handled here (counted, never added to the Undo plan).
        /// </summary>
        private bool TryReencrypt(MailItemRef item, ClassifyRequest request, ClassifyAccumulator acc,
                                  FolderNode sourceFolder, string conversationKey)
        {
            // 1. Only encrypted mail readable decrypted in memory qualifies. The
            //    other verdicts are expected cases, not fallbacks.
            MailProtectionState state;
            try { state = _store.GetProtectionState(item); }
            catch { state = null; }
            var ineligible = ReencryptEligibility.Evaluate(state);
            if (ineligible != null)
            {
                _log.Debug("Re-encrypted filing not applicable: {Reason} ({EntryId})", ineligible.Value, item.EntryId);
                return false;
            }

            // 2. The certificate, resolved once per user action.
            var session = acc.Reencrypt;
            if (!session.CertificateResolved)
            {
                session.CertificateResolved = true;
                try { session.Certificate = _rebuilder.ResolveCertificate(request.ReencryptCertificateThumbprint); }
                catch { session.Certificate = null; }
            }
            var certificate = session.Certificate;
            if (certificate == null)
            {
                if (!_noCertificateReported)
                {
                    _noCertificateReported = true;
                    acc.ReencryptNoCertificate = true;
                }
                return FallBack(item, ReencryptOutcome.NoCertificate, acc);
            }

            // 3. The content as decrypted in memory.
            MailContent content;
            try { content = _rebuilder.ReadDecryptedContent(item); }
            catch { content = null; }
            if (content == null) return FallBack(item, ReencryptOutcome.Unsupported, acc);

            // 4. This item's attachment rows, and its "Save to" files written from
            //    the source (readable in memory) before anything is removed. A
            //    failed save keeps every attachment: the strip is not planned.
            var rows = request.AttachmentDispositions
                .Where(d => d.Item.StoreId == item.StoreId && d.Item.EntryId == item.EntryId)
                .ToList();
            bool removeAttachments = request.RemoveAttachments;
            if (request.RemoveAttachments || request.AttachmentDispositions.Any(row => row.IsInlineImage))
            {
                bool allSavesSucceeded = true;
                foreach (var d in rows)
                {
                    if (d.Action != AttachmentDispositionAction.SaveTo || string.IsNullOrWhiteSpace(d.TargetDirectory))
                        continue;
                    bool saved;
                    try { saved = _store.SaveAttachmentToFile(item, d.AttachmentId, d.TargetDirectory); }
                    catch (Exception ex)
                    {
                        saved = false;
                        _log.Warning(ex, "Save-to-file failed for attachment {AttachmentId} ({EntryId})",
                            d.AttachmentId, item.EntryId);
                    }
                    if (!saved) allSavesSucceeded = false;
                }
                if (!allSavesSucceeded)
                {
                    _log.Warning("Attachment removal skipped: prior save failed ({EntryId})", item.EntryId);
                    acc.Errors++;
                    removeAttachments = false;
                    rows.Clear();
                }
            }
            else
            {
                rows.Clear();
            }

            // 5. Plan the rebuilt content and build its MIME entity.
            byte[] mime;
            int partCount;
            try
            {
                var plan = ReencryptContentPlanner.Plan(content, removeAttachments, rows,
                    request.StripBanner ? request.BannerSignature : null, request.LabelOptions, DateTime.Now);
                if (plan.AttachmentMapMismatch)
                    return FallBack(item, ReencryptOutcome.Unsupported, acc);
                partCount = plan.Content.Parts.Count;
                mime = new MimeEntityBuilder().Build(plan.Content);
            }
            catch (Exception ex)
            {
                // OutOfMemoryException included: a large mail in a 32-bit Outlook.
                _log.Warning(ex, "Re-encrypted filing: building the content failed ({EntryId})", item.EntryId);
                return FallBack(item, ReencryptOutcome.BuildFailed, acc);
            }
            content = null; // the plaintext is no longer needed once built

            // 6. Encrypt, then wipe the plaintext entity.
            byte[] payload;
            try { payload = _rebuilder.Encrypt(mime, certificate.Thumbprint); }
            catch { payload = null; }
            finally { Array.Clear(mime, 0, mime.Length); }
            if (payload == null) return FallBack(item, ReencryptOutcome.EncryptFailed, acc);

            // 7. One verified S/MIME copy per destination; on any failure remove
            //    every copy made for this item and take the standard path.
            var copies = new List<MailItemRef>();
            foreach (var destination in request.Destinations)
            {
                MailItemRef copy;
                try { copy = _rebuilder.CreateEncryptedCopy(item, destination, payload, request.MarkAsRead); }
                catch { copy = null; }
                if (copy == null)
                {
                    DeleteCopies(copies);
                    return FallBack(item, ReencryptOutcome.CopyFailed, acc);
                }
                copies.Add(copy);

                bool verified;
                try { verified = _rebuilder.VerifyEncryptedCopy(copy); }
                catch { verified = false; }
                if (!verified)
                {
                    DeleteCopies(copies);
                    return FallBack(item, ReencryptOutcome.VerifyFailed, acc);
                }
            }

            // 8. The original goes to Deleted Items unless "keep a copy" is on.
            if (!request.KeepCopy)
                MoveOriginalToDeletedItems(item, sourceFolder, acc);

            acc.Processed++;
            acc.Reencrypted++;
            if (!string.IsNullOrEmpty(conversationKey))
                acc.ReencryptedConversationKeys.Add(conversationKey);
            _log.Information(
                "Re-encrypted filing: {Destinations} destination(s), {Parts} part(s), cert {Thumbprint} ({EntryId})",
                copies.Count, partCount, certificate.Thumbprint, item.EntryId);
            return true;
        }

        private bool FallBack(MailItemRef item, ReencryptOutcome reason, ClassifyAccumulator acc)
        {
            acc.ReencryptFallbacks++;
            _log.Warning("Re-encrypted filing fell back to the standard path: {Reason} ({EntryId})",
                reason, item.EntryId);
            return false;
        }

        private void DeleteCopies(IEnumerable<MailItemRef> copies)
        {
            foreach (var copy in copies)
            {
                try { _store.DeleteItemPermanently(copy); }
                catch (Exception ex)
                {
                    _log.Warning(ex, "Re-encrypted filing: removing a partial copy failed ({EntryId})", copy.EntryId);
                }
            }
        }

        /// <summary>
        /// Move the re-encrypted original to its store's Deleted Items. A move can
        /// write the in-memory decrypted state (#62): when the moved item no longer
        /// reads as protected it is deleted for good, so no plaintext stays. A
        /// failure leaves the original where it is; the filing itself stands.
        /// </summary>
        private void MoveOriginalToDeletedItems(MailItemRef item, FolderNode sourceFolder, ClassifyAccumulator acc)
        {
            var trash = ResolveDeletedItems(item.StoreId, ref acc.DeletedItemsByStore);
            if (trash == null)
            {
                _log.Warning("Re-encrypted filing: no Deleted Items folder, original left in place ({EntryId})", item.EntryId);
                acc.Errors++;
                return;
            }

            // Already in Deleted Items: Outlook refuses a same-folder move.
            if (sourceFolder != null && sourceFolder.StoreId == trash.StoreId && sourceFolder.EntryId == trash.EntryId)
                return;

            MailItemRef moved;
            try { moved = _store.MoveItemToFolder(item, trash, false); }
            catch (Exception ex)
            {
                _log.Warning(ex, "Re-encrypted filing: moving the original to Deleted Items failed ({EntryId})", item.EntryId);
                moved = null;
            }
            if (moved == null)
            {
                acc.Errors++;
                return;
            }

            bool stillProtected;
            try { stillProtected = _store.IsEncryptedMail(moved); }
            catch
            {
                _log.Warning("Re-encrypted filing: the original in Deleted Items could not be checked ({EntryId})", moved.EntryId);
                return;
            }
            if (stillProtected) return;

            try
            {
                // Delete() inside Deleted Items is a hard delete.
                _store.DeleteItem(moved);
                _log.Warning("Re-encrypted filing: the original was written decrypted by the move and was deleted " +
                             "permanently ({EntryId})", moved.EntryId);
            }
            catch (Exception ex)
            {
                _log.Warning(ex, "Re-encrypted filing: the original was written decrypted by the move and could not " +
                                 "be deleted ({EntryId})", moved.EntryId);
                acc.Errors++;
            }
        }

        /// <summary>
        /// Remove attachments from the filed item (v2.4.0.0 F2). With per-attachment
        /// dispositions (Modal mode), first save every "Save to" attachment into its
        /// chosen directory, then strip; without them (DeleteSilently, or no rows for
        /// this item), strip all - the pre-F2 behaviour. Acts on
        /// <paramref name="filedRef"/> (the filed copy / moved item), never the kept
        /// original. <see cref="IMailStore.RemoveAttachments"/> stays S/MIME-safe.
        /// </summary>
        private void ApplyAttachmentRemoval(MailItemRef originalItem, MailItemRef filedRef,
                                            ClassifyRequest request, ClassifyAccumulator acc)
        {
            var rows = new List<AttachmentDisposition>();
            foreach (var d in request.AttachmentDispositions)
            {
                if (d.Item.StoreId == originalItem.StoreId && d.Item.EntryId == originalItem.EntryId)
                    rows.Add(d);
            }

            // Save the "Save to" attachments before stripping, while they still
            // exist on the filed item.
            bool allSavesSucceeded = true;
            foreach (var d in rows)
            {
                if (d.Action == AttachmentDispositionAction.SaveTo
                    && !string.IsNullOrWhiteSpace(d.TargetDirectory))
                {
                    bool saved;
                    try { saved = _store.SaveAttachmentToFile(filedRef, d.AttachmentId, d.TargetDirectory); }
                    catch (Exception ex)
                    {
                        saved = false;
                        _log.Warning(ex, "Save-to-file failed for attachment {AttachmentId} on {Subject}",
                            d.AttachmentId, originalItem.Subject);
                    }
                    if (!saved) allSavesSucceeded = false;
                }
            }

            // Never strip when a requested save failed: the only strip primitive
            // removes ALL attachments, so stripping after a failed save would
            // lose the very file the user asked to keep. Leave every attachment
            // in place and count it so the user is told to retry.
            if (rows.Count > 0 && !allSavesSucceeded)
            {
                _log.Warning("Attachment removal skipped: prior save failed for {Subject}", originalItem.Subject);
                acc.Errors++;
                return;
            }

            // Body images (v2.11.0.0 C2) are removed by rewriting the body, never
            // by index or by strip-all, which leaves inline parts alone.
            var regular = rows.Where(d => !d.IsInlineImage).ToList();
            var inline = rows.Where(d => d.IsInlineImage).ToList();

            // B1: a "Keep" row leaves its attachment on the filed copy. When any
            // Keep is present, the strip-all primitive would also take the kept
            // ones, so remove only the non-Keep attachments individually (highest
            // id first, since removing one shifts the indexes of those above it).
            // With no Keep rows the single strip-all call is kept (also the
            // DeleteSilently path, where there are no rows at all). Only body-image
            // rows means there is no regular attachment to strip.
            bool stripped = false;
            if (request.RemoveAttachments && (rows.Count == 0 || regular.Count > 0))
            {
                if (regular.Any(d => d.Action == AttachmentDispositionAction.Keep))
                {
                    foreach (var d in regular.Where(d => d.Action != AttachmentDispositionAction.Keep)
                                             .OrderByDescending(d => d.AttachmentId))
                    {
                        bool removed;
                        try { removed = _store.RemoveAttachment(filedRef, d.AttachmentId); }
                        catch (Exception ex)
                        {
                            removed = false;
                            _log.Warning(ex, "Attachment {AttachmentId} removal failed for {Subject}",
                                d.AttachmentId, originalItem.Subject);
                        }
                        if (removed) stripped = true;
                    }
                }
                else
                {
                    try
                    {
                        stripped = _store.RemoveAttachments(filedRef);
                        // false = the adapter's S/MIME guard fired, not a COM fault.
                        if (!stripped) acc.EncryptedSkips++;
                    }
                    catch
                    {
                        // COM fault during strip-all: the adapter logs the cause.
                        // The filing itself succeeded - propagating would let the
                        // per-destination catch double-count this action as both
                        // filed and errored (v2.7.1.0 B2).
                        stripped = false;
                    }
                }
            }

            // F3: record only the attachments that left the filed copy (Saved /
            // Deleted). Kept attachments are still on the mail, so they need no
            // note (v2.5.0.0 B1).
            var labelRows = stripped
                ? regular.Where(d => d.Action != AttachmentDispositionAction.Keep).ToList()
                : new List<AttachmentDisposition>();

            bool imagesRemoved = RemoveInlineImages(filedRef, inline,
                d => d.Action != AttachmentDispositionAction.Keep, labelRows, originalItem.Subject);

            if (stripped || imagesRemoved) acc.Strips++;

            if (request.LabelOptions != null && labelRows.Count > 0)
            {
                var block = AttachmentLabelFormatter.Format(labelRows, request.LabelOptions, DateTime.Now);
                if (block != null)
                {
                    try { _store.AppendHtmlNote(filedRef, block); }
                    catch { /* best-effort; a label write never fails the filing */ }
                }
            }
        }

        /// <summary>
        /// Remove the body images whose rows <paramref name="isRemoved"/> selects
        /// (v2.11.0.0 C2), in one store call. An image also referenced by a row
        /// that stays is kept. The removed rows are added to
        /// <paramref name="labelRows"/> only when the removal succeeded.
        /// </summary>
        private bool RemoveInlineImages(MailItemRef target, IReadOnlyList<AttachmentDisposition> inline,
                                        Func<AttachmentDisposition, bool> isRemoved,
                                        List<AttachmentDisposition> labelRows, string subject)
        {
            var kept = new HashSet<string>(inline.Where(d => !isRemoved(d)).Select(d => d.ContentId),
                                           StringComparer.OrdinalIgnoreCase);
            var removedRows = inline.Where(d => isRemoved(d) && !kept.Contains(d.ContentId)).ToList();
            if (removedRows.Count == 0) return false;

            var contentIds = removedRows.Select(d => d.ContentId)
                                        .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            bool removed;
            try { removed = _store.RemoveInlineImages(target, contentIds, removedRows[0].InlinePlaceholder); }
            catch (Exception ex)
            {
                removed = false;
                _log.Warning(ex, "Body image removal failed for {Subject}", subject);
            }
            if (removed) labelRows.AddRange(removedRows);
            return removed;
        }

        /// <summary>
        /// A2 (v2.8.0.0): apply attachment dispositions in place on the selected
        /// items, with no classify move. Same rules as
        /// <see cref="ApplyAttachmentRemoval"/> but the item is its own "filed"
        /// ref (nothing was copied/moved). Each item is processed independently.
        /// </summary>
        public AttachmentsAppliedResult ApplyAttachmentDispositions(
            IReadOnlyList<AttachmentDisposition> dispositions,
            AttachmentLabelOptions labelOptions = null,
            bool stripSavedAttachments = true)
        {
            var byItem = new Dictionary<string, KeyValuePair<MailItemRef, List<AttachmentDisposition>>>();
            foreach (var d in dispositions ?? new AttachmentDisposition[0])
            {
                string key = d.Item.StoreId + " " + d.Item.EntryId;
                if (!byItem.TryGetValue(key, out var entry))
                {
                    entry = new KeyValuePair<MailItemRef, List<AttachmentDisposition>>(
                        d.Item, new List<AttachmentDisposition>());
                    byItem[key] = entry;
                }
                entry.Value.Add(d);
            }

            int applied = 0, errors = 0, encryptedSkips = 0;

            foreach (var kv in byItem)
            {
                var item = kv.Value.Key;
                var rows = kv.Value.Value;

                // Save the "Save to" attachments first; block the strip if any fail
                // (the strip primitive removes all attachments, so a failed save
                // would lose the very file the user asked to keep).
                bool anySaved = false;
                bool allSavesSucceeded = true;
                foreach (var d in rows)
                {
                    if (d.Action == AttachmentDispositionAction.SaveTo
                        && !string.IsNullOrWhiteSpace(d.TargetDirectory))
                    {
                        bool saved;
                        try { saved = _store.SaveAttachmentToFile(item, d.AttachmentId, d.TargetDirectory); }
                        catch { saved = false; }
                        if (saved) anySaved = true; else allSavesSucceeded = false;
                    }
                }

                if (!allSavesSucceeded) { errors++; continue; }

                // A3: a saved (SaveTo) attachment is removed only when the caller
                // asked to strip saved attachments; otherwise it is left on the mail
                // (treated like Keep). Delete rows are always removed.
                Func<AttachmentDisposition, bool> isRemoved = d =>
                    d.Action == AttachmentDispositionAction.Delete
                    || (d.Action == AttachmentDispositionAction.SaveTo && stripSavedAttachments);

                // Body images (v2.11.0.0 C2) go through the body rewrite below.
                var regular = rows.Where(d => !d.IsInlineImage).ToList();
                var toRemove = regular.Where(isRemoved).ToList();
                bool anyKept = regular.Any(d => !isRemoved(d));

                bool stripped = false;
                if (toRemove.Count > 0)
                {
                    if (anyKept)
                    {
                        // Some attachments stay: remove only the targeted ones,
                        // highest id first (removing one shifts the indexes above it).
                        foreach (var d in toRemove.OrderByDescending(d => d.AttachmentId))
                        {
                            bool removed;
                            try { removed = _store.RemoveAttachment(item, d.AttachmentId); }
                            catch { removed = false; }
                            if (removed) stripped = true;
                        }
                    }
                    else
                    {
                        try
                        {
                            stripped = _store.RemoveAttachments(item);
                            if (!stripped) encryptedSkips++; // S/MIME guard fired
                        }
                        catch { stripped = false; }
                    }
                }

                var labelRows = stripped ? toRemove : new List<AttachmentDisposition>();
                bool imagesRemoved = RemoveInlineImages(item, rows.Where(d => d.IsInlineImage).ToList(),
                    isRemoved, labelRows, item.Subject);

                if (stripped || anySaved || imagesRemoved) applied++;

                if (labelOptions != null && labelRows.Count > 0)
                {
                    var block = AttachmentLabelFormatter.Format(labelRows, labelOptions, DateTime.Now);
                    if (block != null)
                    {
                        try { _store.AppendHtmlNote(item, block); }
                        catch { /* best-effort */ }
                    }
                }
            }

            return new AttachmentsAppliedResult(applied, errors, encryptedSkips);
        }

        /// <summary>
        /// Record the history for Auto-class: one batch for the whole classify
        /// action, one row per (filed conversation, destination). Best-effort -
        /// a history write never fails the classify. Returns the batch id(s).
        /// </summary>
        private List<string> WriteHistory(ClassifyRequest request, ClassifyAccumulator acc)
        {
            var historyBatchIds = new List<string>();
            if (request.RecordHistory && _history != null && acc.FiledConversationKeys.Count > 0)
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

            // Re-encrypted filings go in a batch of their own that is NOT returned:
            // they cannot be undone, so their history must not make Undo look
            // available (or be rolled back by an Undo of the plain items).
            if (request.RecordHistory && _history != null && acc.ReencryptedConversationKeys.Count > 0)
            {
                try
                {
                    string batchId = Guid.NewGuid().ToString("N");
                    var whenUtc = DateTime.UtcNow;
                    foreach (var key in acc.ReencryptedConversationKeys)
                        _history.Record(batchId, key, request.Destinations, whenUtc);
                }
                catch { /* Auto-class just won't learn this one */ }
            }
            return historyBatchIds;
        }

        private static ClassifyResult BuildResult(ClassifyAccumulator acc, List<string> historyBatchIds)
        {
            var plan = new ClassifyUndoPlan(acc.Moves, acc.CreatedCopies, acc.Flags, acc.Strips, historyBatchIds);
            return new ClassifyResult(acc.Processed, acc.Copies, acc.Moved, acc.Errors, acc.EncryptedSkips,
                                      plan.IsEmpty ? null : plan,
                                      acc.Reencrypted, acc.ReencryptFallbacks, acc.ReencryptNoCertificate);
        }

        /// <summary>Mutable per-call scratch shared by the sync and async classify paths.</summary>
        private sealed class ClassifyAccumulator
        {
            public ClassifyAccumulator(ReencryptSession session) { Reencrypt = session; }

            public int Processed, Copies, Moved, Errors, EncryptedSkips, Strips;

            // Experimental re-encrypted filing (v2.12.0.0).
            public int Reencrypted, ReencryptFallbacks;
            public bool ReencryptNoCertificate;
            public readonly ReencryptSession Reencrypt;
            public readonly HashSet<string> ReencryptedConversationKeys = new HashSet<string>();

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

        /// <summary>
        /// Re-encryption state shared by one user action (a classify, or every
        /// per-item run of one Auto-class): the certificate is resolved once.
        /// </summary>
        private sealed class ReencryptSession
        {
            public bool CertificateResolved;
            public CertificateInfo Certificate;
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
                catch (Exception ex)
                {
                    errors++;
                    _log.Warning(ex, "Undo: flag-restore failed for {Subject}", flagged.Subject);
                }
            }

            // Copies the classify created are removed outright ("undo
            // completely" - they were never user data, just our duplicates).
            foreach (var copy in plan.CreatedCopies)
            {
                try { _store.DeleteItemPermanently(copy); copiesDeleted++; }
                catch (Exception ex)
                {
                    errors++;
                    _log.Warning(ex, "Undo: copy-delete failed for {Subject}", copy.Subject);
                }
            }

            // Finally put each moved original back where it came from, keeping the
            // post-move reference paired with its folder so the moved-back mail can
            // be re-selected in the explorer below (v2.7.0.0 A1).
            var restored = new List<KeyValuePair<FolderNode, MailItemRef>>();
            foreach (var move in plan.Moves)
            {
                try
                {
                    var back = _store.MoveItemToFolder(move.Current, move.SourceFolder);
                    if (back != null)
                    {
                        movesRestored++;
                        restored.Add(new KeyValuePair<FolderNode, MailItemRef>(move.SourceFolder, back));
                    }
                    else
                        errors++;
                }
                catch (Exception ex)
                {
                    errors++;
                    _log.Warning(ex, "Undo: move-restore failed for {Subject} -> {Path}",
                        move.Current.Subject, move.SourceFolder.FullPath);
                }
            }

            // Roll back what Auto-class learned from this classify (best-effort:
            // stale history would only mis-suggest, never lose data).
            if (_history != null && plan.HistoryBatchIds.Count > 0)
            {
                try { _history.DeleteBatches(plan.HistoryBatchIds); }
                catch { /* leave the history rows; they can be re-classified over */ }
            }

            // Report the moved-back mail so the caller can re-select it in Outlook,
            // ready to re-file elsewhere (v2.7.0.0 A1). The explorer displays one
            // folder at a time, so report the items sharing the first restored
            // item's source folder (the common case: one classify of a
            // multi-selection restores every item to the same folder). The caller
            // defers the actual selection until Outlook's view has refreshed.
            FolderNode restoredFolder = null;
            IReadOnlyList<MailItemRef> restoredItems = new MailItemRef[0];
            if (restored.Count > 0)
            {
                restoredFolder = restored[0].Key;
                restoredItems = restored
                    .Where(p => p.Key.StoreId == restoredFolder.StoreId && p.Key.EntryId == restoredFolder.EntryId)
                    .Select(p => p.Value)
                    .ToList();
            }

            return new UndoResult(movesRestored, copiesDeleted, flagsRestored, errors,
                                  restoredFolder, restoredItems);
        }

        public AutoClassifyResult AutoClassify(IReadOnlyList<MailItemRef> items,
                                               Func<string, string, FolderNode> resolveLiveFolder,
                                               ClassifyOptions options = null,
                                               IReadOnlyList<AttachmentDisposition> attachmentDispositions = null,
                                               AttachmentLabelOptions labelOptions = null,
                                               int historyRetentionDays = 0,
                                               string bannerSignature = null)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (resolveLiveFolder == null) throw new ArgumentNullException(nameof(resolveLiveFolder));
            options = options ?? ClassifyOptions.Default;
            if (_history == null)
                throw new InvalidOperationException("Auto-class requires a classification history store.");

            int filed = 0, noHistory = 0, staleFolders = 0, errors = 0,
                partialErrors = 0, encryptedStripSkips = 0, alreadyInPlace = 0;

            DateTime cutoff = (historyRetentionDays > 0)
                ? DateTime.UtcNow.AddDays(-historyRetentionDays)
                : DateTime.MinValue;

            // One combined undo plan over every per-item filing this run does.
            var allMoves = new List<UndoableMove>();
            var allCopies = new List<MailItemRef>();
            var allFlags = new List<MailItemRef>();
            int strips = 0;
            var allBatchIds = new List<string>();

            // Distinct live folders we actually filed into, to show the user.
            var filedDestinations = new List<FolderNode>();
            var filedDestKeys = new HashSet<string>();

            int reencrypted = 0, reencryptFallbacks = 0;
            bool reencryptNoCertificate = false;
            var reencryptSession = new ReencryptSession();

            foreach (var item in Dedupe(items))
            {
                string key;
                try { key = _store.GetConversationKey(item); }
                catch { key = null; }

                if (string.IsNullOrEmpty(key)) { noHistory++; continue; }

                var recorded = _history.GetLatestDestinations(key, cutoff);
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

                // Drop destinations the mail already sits in: Outlook rejects a
                // same-folder move (0x80020009) and there is nothing to do anyway.
                // A mail whose only remembered folder is its current one is
                // "already in place", not an error (v2.7.1.0 B3).
                FolderNode parent;
                try { parent = _store.GetParentFolder(item); }
                catch { parent = null; }
                if (parent != null)
                    live.RemoveAll(f => f.StoreId == parent.StoreId && f.EntryId == parent.EntryId);
                if (live.Count == 0) { alreadyInPlace++; continue; }

                // Reuse the full classify path - same move/copy/strip/history
                // semantics, and it records a fresh batch that Undo rolls back.
                var result = Classify(new ClassifyRequest(
                    new[] { item }, live, options,
                    bannerSignature: bannerSignature,
                    attachmentDispositions: attachmentDispositions,
                    labelOptions: labelOptions), reencryptSession);

                reencrypted += result.Reencrypted;
                reencryptFallbacks += result.ReencryptFallbacks;
                reencryptNoCertificate |= result.ReencryptNoCertificate;

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
                    // Filed, but some destination may still have failed - carry
                    // the partial failures instead of masking them (v2.7.1.0 B3).
                    partialErrors += result.Errors;
                    encryptedStripSkips += result.EncryptedStripSkips;
                }
                else
                {
                    errors++;
                }
            }

            _log.Debug("AutoClassify done: filed={Filed}, noHistory={NoHistory}, staleFolders={StaleFolders}, " +
                       "alreadyInPlace={AlreadyInPlace}, errors={Errors}, partialErrors={PartialErrors}",
                filed, noHistory, staleFolders, alreadyInPlace, errors, partialErrors);

            var plan = new ClassifyUndoPlan(allMoves, allCopies, allFlags, strips, allBatchIds);
            return new AutoClassifyResult(filed, noHistory, staleFolders, errors,
                                          partialErrors, encryptedStripSkips, alreadyInPlace,
                                          plan.IsEmpty ? null : plan, filedDestinations,
                                          reencrypted, reencryptFallbacks, reencryptNoCertificate);
        }

        /// <summary>
        /// C1 (v2.8.0.0): the history/live-folder/already-in-place scan of
        /// <see cref="AutoClassify"/> without any filing, so the pane can decide
        /// which items are processable before it opens the attachment modal.
        /// </summary>
        public AutoClassPreflightResult PreflightAutoClass(
            IReadOnlyList<MailItemRef> items,
            Func<string, string, FolderNode> resolveLiveFolder,
            int historyRetentionDays = 0)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (resolveLiveFolder == null) throw new ArgumentNullException(nameof(resolveLiveFolder));
            if (_history == null)
                throw new InvalidOperationException("Auto-class preflight requires a classification history store.");

            int noHistory = 0, staleFolders = 0, alreadyInPlace = 0;
            DateTime cutoff = (historyRetentionDays > 0)
                ? DateTime.UtcNow.AddDays(-historyRetentionDays)
                : DateTime.MinValue;

            var processable = new List<MailItemRef>();

            foreach (var item in Dedupe(items))
            {
                string key;
                try { key = _store.GetConversationKey(item); }
                catch { key = null; }

                if (string.IsNullOrEmpty(key)) { noHistory++; continue; }

                var recorded = _history.GetLatestDestinations(key, cutoff);
                if (recorded.Count == 0) { noHistory++; continue; }

                var live = new List<FolderNode>();
                foreach (var dest in recorded)
                {
                    FolderNode node;
                    try { node = resolveLiveFolder(dest.StoreId, dest.EntryId); }
                    catch { node = null; }
                    if (node != null) live.Add(node);
                }

                if (live.Count == 0) { staleFolders++; continue; }

                FolderNode parent;
                try { parent = _store.GetParentFolder(item); }
                catch { parent = null; }
                if (parent != null)
                    live.RemoveAll(f => f.StoreId == parent.StoreId && f.EntryId == parent.EntryId);
                if (live.Count == 0) { alreadyInPlace++; continue; }

                processable.Add(item);
            }

            return new AutoClassPreflightResult(processable, noHistory, staleFolders, alreadyInPlace);
        }

        /// <summary>
        /// The original items plus every conversation sibling reported by the
        /// store (legacy 5b step 2), deduped by (StoreId, EntryId) - the
        /// originals win the slot when a sibling lookup also returns one of
        /// them back.
        /// </summary>
        private IReadOnlyList<MailItemRef> Widen(IReadOnlyList<MailItemRef> items,
                                                 List<string> skippedEncrypted,
                                                 out bool wideningTableDisabled)
        {
            var seen = new HashSet<(string StoreId, string EntryId)>();
            var result = new List<MailItemRef>();
            wideningTableDisabled = false;

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
                ConversationSiblings raw;
                try { raw = _store.GetConversationSiblings(item); }
                catch { raw = null; } // adapter failure tolerated; primary item(s) still file
                var siblings = raw ?? ConversationSiblings.Empty;
                if (siblings.WideningTableDisabledThisSession) wideningTableDisabled = true;
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

        private bool SafeIsEncrypted(MailItemRef item)
        {
            try { return _store.IsEncryptedMail(item); }
            catch { return true; } // unknown counts as protected: skip the Save()
        }

        /// <summary>
        /// Record an item's message class and security flags, so a target run can
        /// show whether classify changed them (v2.11.0.0 A3, #62). Reads COM only
        /// when Debug logging is on, and a failure here never affects the filing.
        /// Logs no subject, body or recipient.
        /// </summary>
        private MailProtectionState TraceProtection(string stage, string path, MailItemRef item)
        {
            if (!_log.IsDebugEnabled) return null;

            try
            {
                var state = _store.GetProtectionState(item);
                string messageClass = state == null ? "unresolved" : state.MessageClass ?? "unreadable";
                string flags = state?.SecurityFlags == null
                    ? "absent"
                    : "0x" + state.SecurityFlags.Value.ToString("X8");
                string stored = state == null ? "unresolved" : state.StoredMessageClass ?? "not-found";
                _log.Debug(
                    "Protection trace: stage={Stage} path={Path} entry={EntryId} class={MessageClass} flags={SecurityFlags} " +
                    "protected={Protected} stored={StoredClass} payload={Payload} unsaved={Unsaved} reason={Reason}",
                    stage, path, item.EntryId, messageClass, flags, state?.IsProtected,
                    stored, state?.HasSmimePayload, state?.HasUnsavedChanges, state?.Reason);
                return state;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// A mail protected before filing whose filed item no longer reads as
        /// protected was written decrypted by the move/copy (#62: Stormshield had
        /// decrypted it in memory). RBLclass cannot undo that; the warning makes it
        /// visible in the log. Logs no subject.
        /// </summary>
        private void WarnIfProtectionLost(MailItemRef filedRef, string path)
        {
            bool stillProtected;
            try { stillProtected = _store.IsEncryptedMail(filedRef); }
            catch { return; } // unknown: nothing to report
            if (!stillProtected)
                _log.Warning(
                    "Protected mail filed by {Path} no longer reads as protected ({EntryId}): " +
                    "it was decrypted in memory when filed. No change was made to it.",
                    path, filedRef.EntryId);
        }

        private void TraceProtectionAfter(MailProtectionState before, string path, MailItemRef item)
        {
            var after = TraceProtection("after", path, item);
            if (before != null && before.IsProtected && after != null && !after.IsProtected)
                _log.Warning("Protection trace: {Path} lost protection on {EntryId} ({Before} -> {After})",
                    path, item.EntryId, before.MessageClass, after.MessageClass);
        }
    }
}
