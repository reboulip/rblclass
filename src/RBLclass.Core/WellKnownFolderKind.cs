namespace RBLclass.Core
{
    /// <summary>
    /// Well-known folder roles the Adapter can recognise from the Outlook OM
    /// and pass to <see cref="FolderExclusionPolicy"/>. Keeping this an enum
    /// (rather than matching folder-name strings) is the locale-free part of
    /// deviation #8.
    /// </summary>
    public enum WellKnownFolderKind
    {
        /// <summary>An ordinary user folder.</summary>
        Normal = 0,

        /// <summary>The store's Deleted Items folder.</summary>
        DeletedItems = 1
    }
}
