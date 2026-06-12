namespace RBLclass.Core
{
    /// <summary>
    /// How a keyword is tested against a folder path in <see cref="IFolderSearch"/>.
    /// Exposed so it can become a user preference (Step 9 settings).
    /// </summary>
    public enum FolderMatchMode
    {
        /// <summary>
        /// Opt-in stricter mode. Each keyword must be the prefix of some word in
        /// the path (words split on non-alphanumeric; case- and accent-
        /// insensitive). e.g. "proj" matches "Projects" but "security" does NOT
        /// match "Cybersecurity". Less noisy than substring.
        /// </summary>
        WordPrefix = 0,

        /// <summary>
        /// Default. Each keyword may appear anywhere in the path (spaces removed;
        /// case- and accent-insensitive). e.g. "security" matches "Cybersecurity"
        /// and "juri" matches inside "ProjetJuridique". Broadest, most forgiving.
        /// </summary>
        Substring = 1
    }
}
