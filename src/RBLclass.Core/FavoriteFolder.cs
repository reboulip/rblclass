using System;

namespace RBLclass.Core
{
    /// <summary>
    /// One indexed favourite directory: a Windows filesystem path the user can
    /// save attachments into, searchable like an Outlook folder (v2.4.0.0 F1).
    /// Paths only - never file contents (32-bit memory budget).
    /// </summary>
    public sealed class FavoriteFolder
    {
        public FavoriteFolder(string path, string displayName)
        {
            Path = path ?? throw new ArgumentNullException(nameof(path));
            DisplayName = displayName ?? throw new ArgumentNullException(nameof(displayName));
        }

        /// <summary>Full filesystem path (the search surface and the save target).</summary>
        public string Path { get; }

        /// <summary>Last path segment, shown as the folder's name.</summary>
        public string DisplayName { get; }
    }
}
