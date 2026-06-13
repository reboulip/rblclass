using System;
using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// One move performed by a classify, recorded so it can be reversed:
    /// the filed item (by its post-move reference) and the folder it came from.
    /// </summary>
    public sealed class UndoableMove
    {
        public UndoableMove(MailItemRef current, FolderNode sourceFolder)
        {
            Current = current ?? throw new ArgumentNullException(nameof(current));
            SourceFolder = sourceFolder ?? throw new ArgumentNullException(nameof(sourceFolder));
        }

        /// <summary>The moved item where it now lives (destination store/entry).</summary>
        public MailItemRef Current { get; }

        /// <summary>The folder the item was moved out of.</summary>
        public FolderNode SourceFolder { get; }
    }

    /// <summary>
    /// Everything a classify did that can be reversed (v2.2 Undo): moves to put
    /// back, copies (incl. Deleted Items safety copies) to remove, follow-up
    /// flags to un-complete. Attachment strips are unrecoverable and only
    /// counted so the UI can say so. Single-slot: the pane keeps the latest
    /// plan only.
    /// </summary>
    public sealed class ClassifyUndoPlan
    {
        public ClassifyUndoPlan(IReadOnlyList<UndoableMove> moves,
                                IReadOnlyList<MailItemRef> createdCopies,
                                IReadOnlyList<MailItemRef> completedFlags,
                                int attachmentStrips,
                                IReadOnlyList<string> historyBatchIds = null)
        {
            Moves = (moves ?? new UndoableMove[0]).ToArray();
            CreatedCopies = (createdCopies ?? new MailItemRef[0]).ToArray();
            CompletedFlags = (completedFlags ?? new MailItemRef[0]).ToArray();
            AttachmentStrips = attachmentStrips;
            HistoryBatchIds = (historyBatchIds ?? new string[0]).ToArray();
        }

        /// <summary>Items the classify moved, with where to put them back.</summary>
        public IReadOnlyList<UndoableMove> Moves { get; }

        /// <summary>Copies the classify created - undone by deleting them permanently.</summary>
        public IReadOnlyList<MailItemRef> CreatedCopies { get; }

        /// <summary>Filed items whose follow-up flag was marked complete - undone by re-marking.</summary>
        public IReadOnlyList<MailItemRef> CompletedFlags { get; }

        /// <summary>Filed items whose attachments were stripped - NOT recoverable by Undo.</summary>
        public int AttachmentStrips { get; }

        /// <summary>Classification-history batches this classify recorded - undone by deleting them.</summary>
        public IReadOnlyList<string> HistoryBatchIds { get; }

        public bool IsEmpty =>
            Moves.Count == 0 && CreatedCopies.Count == 0 && CompletedFlags.Count == 0
            && HistoryBatchIds.Count == 0;
    }

    /// <summary>Outcome of executing a <see cref="ClassifyUndoPlan"/>.</summary>
    public sealed class UndoResult
    {
        public UndoResult(int movesRestored, int copiesDeleted, int flagsRestored, int errors)
        {
            MovesRestored = movesRestored;
            CopiesDeleted = copiesDeleted;
            FlagsRestored = flagsRestored;
            Errors = errors;
        }

        public int MovesRestored { get; }
        public int CopiesDeleted { get; }
        public int FlagsRestored { get; }
        public int Errors { get; }
    }
}
