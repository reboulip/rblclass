namespace RBLclass.Core
{
    /// <summary>Observable state of the folder index, driven by IFolderIndexService.</summary>
    public enum IndexStatus
    {
        /// <summary>Index absent or empty — never built, or DB cleared.</summary>
        NotFound = 0,
        /// <summary>A live store walk is currently in progress.</summary>
        Indexing = 1,
        /// <summary>Index is populated and ready for search.</summary>
        Ready = 2
    }
}
