using System;

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
                    foreach (var destination in request.Destinations)
                    {
                        var copy = _store.CopyItemToFolder(item, destination);
                        copies++;

                        // Strip attachments from the FILED COPY only, never the
                        // original - so "keep a copy" leaves the original intact
                        // with its attachments.
                        if (request.RemoveAttachments && copy != null)
                            _store.RemoveAttachments(copy);
                    }

                    if (!request.KeepCopy)
                    {
                        _store.DeleteItem(item);
                        deleted++;
                    }

                    processed++;
                }
                catch
                {
                    // Skip this item, keep going; the adapter/caller logs details.
                    errors++;
                }
            }

            return new ClassifyResult(processed, copies, deleted, errors);
        }
    }
}
