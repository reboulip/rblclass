using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// Persistence seam for the favourite-folder filesystem index (v2.4.0.0 F1),
    /// so the expanded directory tree survives between sessions and search is
    /// fast at startup without re-walking the filesystem. Implemented over SQLite
    /// by <c>RBLclass.Core.Persistence.SqliteFolderRepository</c> (same database
    /// as the Outlook folder index).
    /// </summary>
    public interface IFavoriteFolderRepository
    {
        /// <summary>Replace the entire favourite-folder index in one transaction.</summary>
        void SaveFavorites(IEnumerable<FavoriteFolder> folders);

        /// <summary>Load the full persisted favourite-folder set, ordered by path.</summary>
        IReadOnlyList<FavoriteFolder> LoadFavorites();
    }
}
