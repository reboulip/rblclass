using System;
using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>A past classify destination, as stored in the history (keys only - resolved against the live folder index when used).</summary>
    public sealed class HistoryDestination
    {
        public HistoryDestination(string storeId, string entryId)
        {
            StoreId = storeId ?? throw new ArgumentNullException(nameof(storeId));
            EntryId = entryId ?? throw new ArgumentNullException(nameof(entryId));
        }

        public string StoreId { get; }
        public string EntryId { get; }
    }

    /// <summary>
    /// Where past classify actions are remembered, keyed by Outlook
    /// conversation (v2.2 Auto-class): every successful classify appends one
    /// row per (conversation, destination) under a batch id, and Auto-class
    /// replays the latest batch for a conversation. Undo deletes the batches
    /// a classify wrote. Implemented over SQLite (schema v2) by
    /// <c>RBLclass.Core.Persistence.SqliteClassificationHistory</c>.
    /// </summary>
    public interface IClassificationHistory
    {
        /// <summary>Create the history table if missing (idempotent).</summary>
        void EnsureSchema();

        /// <summary>
        /// Append one history row per destination for a conversation, all under
        /// <paramref name="batchId"/> (one batch = one classify action, so the
        /// "latest action" for a conversation is a whole destination set).
        /// </summary>
        void Record(string batchId, string conversationKey,
                    IEnumerable<FolderNode> destinations, DateTime whenUtc);

        /// <summary>
        /// The destination set of the most recent recorded classify for a
        /// conversation that is not older than <paramref name="notOlderThan"/> (UTC),
        /// or an empty list when the conversation was never classified or all
        /// history is outside the retention window. Pass <see cref="DateTime.MinValue"/>
        /// to disable the window (no filter).
        /// </summary>
        IReadOnlyList<HistoryDestination> GetLatestDestinations(
            string conversationKey, DateTime notOlderThan);

        /// <summary>Remove every row of the given batches (Undo rolling back what a classify learned).</summary>
        void DeleteBatches(IEnumerable<string> batchIds);
    }
}
