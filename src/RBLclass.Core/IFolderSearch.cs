namespace RBLclass.Core
{
    /// <summary>
    /// Keyword search over the cached folder tree. Reproduces the legacy
    /// <c>kwSearch</c> semantics: AND across keywords, case/accent/space
    /// insensitive, with collapse/expand and a result cap. Matches in memory
    /// over the folder set (deviation #2 - NOT FTS5, which would break partial
    /// matches).
    /// </summary>
    public interface IFolderSearch
    {
        /// <summary>
        /// Find folders matching every keyword in <paramref name="query"/>
        /// (whitespace-split). Returns <see cref="FolderSearchOutcome.Empty"/>
        /// for a blank query.
        /// </summary>
        FolderSearchOutcome Search(string query, FolderSearchOptions options = null);
    }
}
