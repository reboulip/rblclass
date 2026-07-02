namespace RBLclass.Core
{
    /// <summary>
    /// How a folder search-result row renders its path (v2.7.0.0 C1). The
    /// Settings dialog picks the default; the per-row ⤢ button cycles
    /// OneLine → Wrapped → PerSegment → OneLine starting from that default.
    /// </summary>
    public enum SearchPathDisplayMode
    {
        /// <summary>A single left-trimmed line (the default).</summary>
        OneLine = 0,

        /// <summary>The full path on one wrapping line, breaking on pane width.</summary>
        Wrapped = 1,

        /// <summary>The full path with one sub-folder segment per line.</summary>
        PerSegment = 2
    }
}
