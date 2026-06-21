using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// Filesystem directory enumeration, injected so <see cref="RBLclass.Core"/>
    /// stays I/O-free (the walk is not Outlook-specific, so it lives in the
    /// add-in shell rather than the Outlook adapter - v2.4.0.0 F1).
    /// </summary>
    public interface IDirectoryScanner
    {
        /// <summary>
        /// Return <paramref name="rootPath"/> itself followed by its
        /// subdirectory paths, up to <paramref name="maxDepth"/> levels deep and
        /// at most <paramref name="maxTotal"/> entries in total. Returns an empty
        /// list when the root does not exist or is not accessible; never throws.
        /// </summary>
        IReadOnlyList<string> GetDirectories(string rootPath, int maxDepth, int maxTotal);
    }
}
