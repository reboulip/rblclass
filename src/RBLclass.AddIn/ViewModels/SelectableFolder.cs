using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// A folder search result with a checkbox state, so destinations can be
    /// toggled with single clicks (no Ctrl) in the Classify list.
    /// </summary>
    public sealed class SelectableFolder : ObservableObject
    {
        private bool _isSelected;
        private bool _isExpanded;

        public SelectableFolder(FolderSearchResult result)
        {
            Result = result;
        }

        public FolderSearchResult Result { get; }
        public FolderNode Folder => Result.Folder;
        public string DisplayPath => Result.DisplayPath;

        /// <summary>Path split into segments for the per-row expanded multi-line view (v2.4 C1).</summary>
        public string[] PathSegments => Result.PathSegments;

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        /// <summary>
        /// Per-row toggle: when on, this single result shows its full path across
        /// multiple lines (one segment per line, no indentation) so a long leaf is
        /// not truncated. Off by default; independent of every other row (v2.4 C1).
        /// </summary>
        public bool IsExpanded
        {
            get => _isExpanded;
            set => SetProperty(ref _isExpanded, value);
        }
    }
}
