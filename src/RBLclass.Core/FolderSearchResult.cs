using System;

namespace RBLclass.Core
{
    /// <summary>
    /// One folder-search hit: the target folder plus how it should be shown.
    /// </summary>
    public sealed class FolderSearchResult
    {
        /// <summary>Suffix appended to a collapsed non-leaf match (legacy "|[...]").</summary>
        public const string CollapsedSuffix = " | [...]";

        public FolderSearchResult(FolderNode folder, bool isCollapsed)
        {
            Folder = folder ?? throw new ArgumentNullException(nameof(folder));
            IsCollapsed = isCollapsed;
            DisplayPath = isCollapsed ? folder.FullPath + CollapsedSuffix : folder.FullPath;
        }

        /// <summary>The matched folder (the navigation/classify target).</summary>
        public FolderNode Folder { get; }

        /// <summary>
        /// True when this is a collapsed non-leaf match standing in for its
        /// whole matching subtree (only happens when All results is off).
        /// </summary>
        public bool IsCollapsed { get; }

        /// <summary>The string to show in the result list.</summary>
        public string DisplayPath { get; }
    }
}
