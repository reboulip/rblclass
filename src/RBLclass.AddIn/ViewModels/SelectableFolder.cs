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

        public SelectableFolder(FolderSearchResult result)
        {
            Result = result;
        }

        public FolderSearchResult Result { get; }
        public FolderNode Folder => Result.Folder;
        public string DisplayPath => Result.DisplayPath;

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }
    }
}
