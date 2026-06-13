namespace RBLclass.Core
{
    /// <summary>Outcome of a classify operation, for the UI status line and logging.</summary>
    public sealed class ClassifyResult
    {
        public ClassifyResult(int itemsProcessed, int copiesMade, int originalsMoved,
                              int errors, int encryptedStripSkips = 0,
                              ClassifyUndoPlan undo = null)
        {
            ItemsProcessed = itemsProcessed;
            CopiesMade = copiesMade;
            OriginalsMoved = originalsMoved;
            Errors = errors;
            EncryptedStripSkips = encryptedStripSkips;
            Undo = undo;
        }

        /// <summary>Items filed without error.</summary>
        public int ItemsProcessed { get; }

        /// <summary>Total copies placed (extra destinations, or all of them when "keep a copy" is on).</summary>
        public int CopiesMade { get; }

        /// <summary>
        /// Originals moved into their (last) destination (0 when "keep a copy"
        /// is on). Since v2.2 classify moves the original instead of deleting
        /// it after copying - see <see cref="IMailStore.MoveItemToFolder"/>.
        /// </summary>
        public int OriginalsMoved { get; }

        /// <summary>Items that failed (skipped, others still processed).</summary>
        public int Errors { get; }

        /// <summary>
        /// Filed items whose attachments were NOT stripped despite "remove
        /// attachments", because they are S/MIME-encrypted/signed (their
        /// attachments are the message itself - never stripped).
        /// </summary>
        public int EncryptedStripSkips { get; }

        /// <summary>
        /// How to reverse this classify (v2.2 Undo), or null when nothing
        /// reversible happened.
        /// </summary>
        public ClassifyUndoPlan Undo { get; }
    }
}
