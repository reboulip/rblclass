using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>One row of the attachment disposition modal, as <see cref="AttachmentBulkRule"/> sees it.</summary>
    public interface IAttachmentRowState
    {
        AttachmentDispositionAction Action { get; }

        /// <summary>True once the user explicitly picked this row's action.</summary>
        bool IsActionUserSet { get; }

        bool IsInlineImage { get; }
        string TargetDirectory { get; }
    }

    /// <summary>
    /// The "same action and folder for all attachments" rule of the disposition
    /// modal (v2.11.0.0 C1). A group choice never changes a pinned row: one the
    /// user explicitly set to Delete or Keep (v2.7.1.0 A2: never flip such a row
    /// to Save to) or an inline image, whose removal is irreversible.
    /// </summary>
    public static class AttachmentBulkRule
    {
        public static bool IsPinned(IAttachmentRowState row) =>
            row.IsInlineImage
            || (row.IsActionUserSet && row.Action != AttachmentDispositionAction.SaveTo);

        public static bool AppliesTo(IAttachmentRowState row) => !IsPinned(row);

        /// <summary>
        /// The directory a row gets from a group choice: the group's directory for
        /// Save to, otherwise (or when the group has none yet) the row's own.
        /// </summary>
        public static string DirectoryFor(IAttachmentRowState row, AttachmentDispositionAction action, string directory) =>
            action == AttachmentDispositionAction.SaveTo && !string.IsNullOrWhiteSpace(directory)
                ? directory
                : row.TargetDirectory ?? string.Empty;

        /// <summary>
        /// "Action per attachment, one folder" mode: a Save-to row takes the
        /// shared folder, any other row keeps its own (v2.11.0.0 C1).
        /// </summary>
        public static string SharedFolderFor(IAttachmentRowState row, string directory) =>
            row.Action == AttachmentDispositionAction.SaveTo && !string.IsNullOrWhiteSpace(directory)
                ? directory
                : row.TargetDirectory ?? string.Empty;

        /// <summary>Confirm is allowed once every Save-to row has a directory.</summary>
        public static bool CanConfirm(IEnumerable<IAttachmentRowState> rows) =>
            rows.All(r => r.Action != AttachmentDispositionAction.SaveTo
                          || !string.IsNullOrWhiteSpace(r.TargetDirectory));

        /// <summary>
        /// The group choice to show when the user switches back to the grouped
        /// view, read from the unpinned rows only: Save to if any of them saves,
        /// Delete only if all of them delete (never inferred from a mix), Keep
        /// otherwise. The directory is the first one set among them.
        /// </summary>
        public static AttachmentDispositionAction SeedGroupChoice(IEnumerable<IAttachmentRowState> rows, out string directory)
        {
            var unpinned = rows.Where(AppliesTo).ToList();
            directory = unpinned.Select(r => r.TargetDirectory)
                                .FirstOrDefault(d => !string.IsNullOrWhiteSpace(d)) ?? string.Empty;

            if (unpinned.Count == 0 || unpinned.Any(r => r.Action == AttachmentDispositionAction.SaveTo))
                return AttachmentDispositionAction.SaveTo;
            return unpinned.All(r => r.Action == AttachmentDispositionAction.Delete)
                ? AttachmentDispositionAction.Delete
                : AttachmentDispositionAction.Keep;
        }
    }
}
