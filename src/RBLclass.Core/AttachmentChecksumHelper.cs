using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;

namespace RBLclass.Core
{
    /// <summary>
    /// Streaming SHA-256 content-equality check for deduplication of saved
    /// attachments and mail files (v2.9.0.0 D2). Both files are read via
    /// SHA256.ComputeHash(Stream) so no full file load occurs - safe for the
    /// ~1 GB memory budget of the 32-bit Outlook host.
    /// </summary>
    public static class AttachmentChecksumHelper
    {
        public static bool ContentEquals(
            string path1,
            string path2,
            Func<string, Stream> openStream = null)
        {
            if (path1 == null) throw new ArgumentNullException(nameof(path1));
            if (path2 == null) throw new ArgumentNullException(nameof(path2));
            openStream = openStream ?? (p => File.OpenRead(p));

            byte[] h1, h2;
            using (var sha = SHA256.Create())
            {
                using (var s = openStream(path1)) h1 = sha.ComputeHash(s);
                sha.Initialize();
                using (var s = openStream(path2)) h2 = sha.ComputeHash(s);
            }
            return h1.SequenceEqual(h2);
        }
    }
}
