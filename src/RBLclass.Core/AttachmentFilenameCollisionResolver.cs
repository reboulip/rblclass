using System;
using System.IO;

namespace RBLclass.Core
{
    /// <summary>
    /// Resolves a non-colliding destination path for a saved attachment
    /// (v2.4.0.0 F2): if the proposed file already exists, append " (1)", " (2)",
    /// … before the extension until the path is free. Pure - the existence check
    /// is injected so it is unit-testable without touching the filesystem (the
    /// adapter passes <see cref="File.Exists"/>).
    /// </summary>
    public static class AttachmentFilenameCollisionResolver
    {
        public static string Resolve(string directory, string fileName, Func<string, bool> exists)
        {
            if (exists == null) throw new ArgumentNullException(nameof(exists));

            string candidate = Path.Combine(directory, fileName);
            if (!exists(candidate)) return candidate;

            string stem = Path.GetFileNameWithoutExtension(fileName);
            string ext = Path.GetExtension(fileName);
            for (int i = 1; i < 10000; i++)
            {
                candidate = Path.Combine(directory, stem + " (" + i + ")" + ext);
                if (!exists(candidate)) return candidate;
            }
            // Unreachable in practice; fall back to a timestamp suffix.
            return Path.Combine(directory,
                stem + " (" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ")" + ext);
        }
    }
}
