using System.Collections.Generic;
using System.Linq;

namespace RBLclass.Core
{
    /// <summary>
    /// Outcome of an Auto-class run (v2.2): each selected mail is checked
    /// against the classification history and filed to its conversation's most
    /// recent recorded destination(s), validated against the live folder index.
    /// </summary>
    public sealed class AutoClassifyResult
    {
        public AutoClassifyResult(int filed, int noHistory, int staleFolders,
                                  int errors, ClassifyUndoPlan undo,
                                  IReadOnlyList<FolderNode> filedDestinations = null)
        {
            Filed = filed;
            NoHistory = noHistory;
            StaleFolders = staleFolders;
            Errors = errors;
            Undo = undo;
            FiledDestinations = (filedDestinations ?? new FolderNode[0]).ToArray();
        }

        /// <summary>Mails filed to a remembered destination.</summary>
        public int Filed { get; }

        /// <summary>Mails skipped because their conversation has never been classified.</summary>
        public int NoHistory { get; }

        /// <summary>
        /// Mails skipped because every remembered destination no longer exists
        /// in the folder index (the folder was deleted/renamed away).
        /// </summary>
        public int StaleFolders { get; }

        /// <summary>Mails that failed to file despite a valid remembered destination.</summary>
        public int Errors { get; }

        /// <summary>
        /// The combined undo plan covering every filing this run did, or null
        /// when nothing was filed. Undo treats it like any other classify.
        /// </summary>
        public ClassifyUndoPlan Undo { get; }

        /// <summary>
        /// The distinct live folders this run filed mail into, so the pane can
        /// show the user where things went. Empty when nothing was filed.
        /// </summary>
        public IReadOnlyList<FolderNode> FiledDestinations { get; }

        /// <summary>True when at least one mail was filed.</summary>
        public bool AnyFiled => Filed > 0;
    }
}
