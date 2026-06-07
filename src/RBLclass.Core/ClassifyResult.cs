namespace RBLclass.Core
{
    /// <summary>Outcome of a classify operation, for the UI status line and logging.</summary>
    public sealed class ClassifyResult
    {
        public ClassifyResult(int itemsProcessed, int copiesMade, int originalsDeleted, int errors)
        {
            ItemsProcessed = itemsProcessed;
            CopiesMade = copiesMade;
            OriginalsDeleted = originalsDeleted;
            Errors = errors;
        }

        /// <summary>Items filed without error.</summary>
        public int ItemsProcessed { get; }

        /// <summary>Total copies placed (items x destinations).</summary>
        public int CopiesMade { get; }

        /// <summary>Originals deleted (0 when "keep a copy" is on).</summary>
        public int OriginalsDeleted { get; }

        /// <summary>Items that failed (skipped, others still processed).</summary>
        public int Errors { get; }
    }
}
