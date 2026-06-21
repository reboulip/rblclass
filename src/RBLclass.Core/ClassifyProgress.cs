namespace RBLclass.Core
{
    /// <summary>
    /// Per-item progress for a responsive classify (v2.4 D2): how many of the
    /// batch's items have been filed so far. Reported once per processed item by
    /// <see cref="IClassifier.ClassifyAsync"/>.
    /// </summary>
    public sealed class ClassifyProgress
    {
        public ClassifyProgress(int completed, int total)
        {
            Completed = completed;
            Total = total;
        }

        /// <summary>Items filed so far (1-based, incremented as each completes).</summary>
        public int Completed { get; }

        /// <summary>Total items in the batch.</summary>
        public int Total { get; }
    }
}
