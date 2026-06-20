using System.ComponentModel;

namespace RBLclass.Core
{
    /// <summary>
    /// Extends <see cref="IFolderTree"/> with an observable <see cref="IndexStatus"/>
    /// so the add-in shell and pane view model can reflect the walk lifecycle
    /// without knowing the concrete <see cref="FolderIndexService"/>.
    /// </summary>
    public interface IFolderIndexService : IFolderTree, INotifyPropertyChanged
    {
        IndexStatus IndexStatus { get; }
    }
}
