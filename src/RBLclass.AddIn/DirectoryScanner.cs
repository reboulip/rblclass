using System.Collections.Generic;
using System.IO;
using RBLclass.Core;

namespace RBLclass.AddIn
{
    /// <summary>
    /// Shell-layer <see cref="IDirectoryScanner"/> over <see cref="System.IO"/>
    /// (Core does no I/O except via injected interfaces) - v2.4.0.0 F1. Walks a
    /// favourite root to a bounded depth/count, best-effort: an inaccessible
    /// subtree is skipped, never fatal.
    /// </summary>
    internal sealed class DirectoryScanner : IDirectoryScanner
    {
        public IReadOnlyList<string> GetDirectories(string rootPath, int maxDepth, int maxTotal)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(rootPath) || maxTotal <= 0) return result;
            try
            {
                if (!Directory.Exists(rootPath)) return result;
            }
            catch { return result; }

            result.Add(rootPath);
            Recurse(rootPath, 1, maxDepth, maxTotal, result);
            return result;
        }

        private static void Recurse(string dir, int depth, int maxDepth,
                                    int maxTotal, List<string> result)
        {
            if (depth > maxDepth || result.Count >= maxTotal) return;

            string[] children;
            try { children = Directory.GetDirectories(dir); }
            catch { return; } // access denied, junction loop, etc.

            foreach (var child in children)
            {
                if (result.Count >= maxTotal) return;
                result.Add(child);
                Recurse(child, depth + 1, maxDepth, maxTotal, result);
            }
        }
    }
}
