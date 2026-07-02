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
        private SearchPathDisplayMode _displayMode;

        public SelectableFolder(FolderSearchResult result)
            : this(result, SearchPathDisplayMode.OneLine)
        {
        }

        public SelectableFolder(FolderSearchResult result, SearchPathDisplayMode displayMode)
        {
            Result = result;
            _displayMode = displayMode;
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
        /// How this single row renders its path (v2.7.0.0 C1). Independent of every
        /// other row; seeded from the Settings default and cycled by the ⤢ button.
        /// </summary>
        public SearchPathDisplayMode DisplayMode
        {
            get => _displayMode;
            set
            {
                if (SetProperty(ref _displayMode, value))
                {
                    OnPropertyChanged(nameof(IsOneLine));
                    OnPropertyChanged(nameof(IsWrapped));
                    OnPropertyChanged(nameof(IsPerSegment));
                }
            }
        }

        public bool IsOneLine => _displayMode == SearchPathDisplayMode.OneLine;
        public bool IsWrapped => _displayMode == SearchPathDisplayMode.Wrapped;
        public bool IsPerSegment => _displayMode == SearchPathDisplayMode.PerSegment;

        /// <summary>Advance the ⤢ button: OneLine → Wrapped → PerSegment → OneLine (v2.7.0.0 C1).</summary>
        public void CycleDisplayMode()
        {
            DisplayMode = _displayMode == SearchPathDisplayMode.PerSegment
                ? SearchPathDisplayMode.OneLine
                : _displayMode + 1;
        }
    }
}
