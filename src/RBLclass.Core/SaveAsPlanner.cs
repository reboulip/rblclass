using System;
using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// The user's answer in the Save-as picker: the target directory and, when
    /// the format setting is <see cref="SaveAsFormat.Ask"/>, the chosen format.
    /// </summary>
    public sealed class SaveAsChoice
    {
        public SaveAsChoice(string directory, SaveAsFormat format)
        {
            Directory = directory;
            Format = format;
        }

        public string Directory { get; }
        public SaveAsFormat Format { get; }
    }

    /// <summary>A batch split into plain mails and protected (encrypted/signed) ones.</summary>
    public sealed class SaveAsPartition<T>
    {
        public SaveAsPartition(IReadOnlyList<T> plain, IReadOnlyList<T> @protected)
        {
            Plain = plain;
            Protected = @protected;
        }

        public IReadOnlyList<T> Plain { get; }
        public IReadOnlyList<T> Protected { get; }
    }

    /// <summary>
    /// Pure planning for the Save-as action (v2.13.0.0 C1): which format to write,
    /// how a PDF is named, and which mails of a batch need the "encrypted mail
    /// becomes plaintext" confirmation. No Outlook, no I/O except the injected
    /// existence check.
    /// </summary>
    public static class SaveAsPlanner
    {
        /// <summary>
        /// Resolve the concrete format: a fixed setting wins; with
        /// <see cref="SaveAsFormat.Ask"/> the picker's choice decides, and a missing
        /// choice falls back to .msg (the historical format).
        /// </summary>
        public static SaveAsFormat Resolve(SaveAsFormat setting, SaveAsFormat? pickerChoice)
        {
            if (setting == SaveAsFormat.Ask)
                return pickerChoice == SaveAsFormat.Pdf ? SaveAsFormat.Pdf : SaveAsFormat.Msg;
            return setting == SaveAsFormat.Pdf ? SaveAsFormat.Pdf : SaveAsFormat.Msg;
        }

        /// <summary>
        /// PDF destination path: the sanitized subject (the same name the .msg gets,
        /// "email" when blank) with .pdf, suffixed " (1)", " (2)"… on collision. No
        /// checksum dedupe: PDF bytes differ from one run to the next.
        /// </summary>
        public static string PdfPath(string directory, string sanitizedBaseName, Func<string, bool> exists)
        {
            if (directory == null) throw new ArgumentNullException(nameof(directory));
            string baseName = string.IsNullOrWhiteSpace(sanitizedBaseName) ? "email" : sanitizedBaseName;
            return AttachmentFilenameCollisionResolver.Resolve(directory, baseName + ".pdf", exists);
        }

        /// <summary>Split a batch with the shared protected-mail check; order is kept.</summary>
        public static SaveAsPartition<T> Partition<T>(IEnumerable<T> items, Func<T, bool> isProtected)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (isProtected == null) throw new ArgumentNullException(nameof(isProtected));

            var plain = new List<T>();
            var prot = new List<T>();
            foreach (var item in items)
                (isProtected(item) ? prot : plain).Add(item);
            return new SaveAsPartition<T>(plain, prot);
        }
    }
}
