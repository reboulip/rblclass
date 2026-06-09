using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// Persistence seam for the folder index. The folder tree rests here
    /// between sessions, so every start after the first loads from this store
    /// instead of re-walking Outlook (the index lifecycle - see
    /// <see cref="IFolderTree"/>). Implemented over SQLite by
    /// <c>RBLclass.Core.Persistence.SqliteFolderRepository</c>.
    /// </summary>
    public interface IFolderRepository
    {
        /// <summary>
        /// Create/upgrade the schema to the current version (idempotent). Must
        /// be called before any other operation.
        /// </summary>
        void EnsureSchema();

        /// <summary>True if the index already holds at least one folder.</summary>
        bool HasAnyFolders();

        /// <summary>Load the full persisted folder set, ordered by full path.</summary>
        IReadOnlyList<FolderNode> LoadAll();

        /// <summary>Replace the entire folder index in one transaction.</summary>
        void ReplaceAll(IEnumerable<FolderNode> folders);

        /// <summary>
        /// Replace just one store's folders in one transaction (targeted
        /// re-index after a sub-folder change), leaving other stores untouched.
        /// </summary>
        void ReplaceStore(string storeId, IEnumerable<FolderNode> folders);
    }
}
