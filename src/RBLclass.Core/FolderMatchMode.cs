namespace RBLclass.Core
{
    /// <summary>
    /// How a keyword is tested against a folder path in <see cref="IFolderSearch"/>.
    /// Exposed so it can become a user preference (Step 9 settings).
    /// </summary>
    public enum FolderMatchMode
    {
        /// <summary>
        /// Default. Each keyword must be the prefix of some word in the path
        /// (words split on non-alphanumeric; case- and accent-insensitive).
        /// e.g. "proj" matches "Projects". Less noisy than substring.
        /// </summary>
        WordPrefix = 0,

        /// <summary>
        /// Legacy behaviour: each keyword may appear anywhere in the path
        /// (spaces removed; case- and accent-insensitive). e.g. "juri" matches
        /// inside "ProjetJuridique". More matches, more noise - opt-in.
        /// </summary>
        Substring = 1
    }
}
