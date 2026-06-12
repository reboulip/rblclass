using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using RBLclass.AddIn.ViewModels;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// The single unified RBLclass task pane (open + classify in one list).
    /// Thin code-behind: gestures call view-model methods (no MVVM package).
    /// </summary>
    public partial class MainPaneView : UserControl
    {
        public MainPaneView()
        {
            InitializeComponent();
        }

        private MainPaneViewModel Vm => DataContext as MainPaneViewModel;

        private void QueryBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Vm?.FileToHighlighted();
                e.Handled = true;
            }
        }

        private void ResultsList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Ignore double-clicks that land on the per-row buttons.
            if (IsWithinButton(e.OriginalSource as DependencyObject)) return;

            // File into the row that was actually double-clicked, not the list's
            // current selection (clicking a row toggles its checkbox without
            // changing the ListBox selection).
            var folder = ResolveFolder(e.OriginalSource as DependencyObject);
            if (folder != null) Vm?.FileToFolder(folder.Folder);
        }

        private void Classify_Click(object sender, RoutedEventArgs e) => Vm?.ClassifyChecked();

        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            var folder = (sender as FrameworkElement)?.DataContext as SelectableFolder;
            if (folder != null) Vm?.OpenFolder(folder.Folder);
        }

        private void NewSubfolder_Click(object sender, RoutedEventArgs e)
        {
            var folder = (sender as FrameworkElement)?.DataContext as SelectableFolder;
            if (folder != null) Vm?.CreateSubfolderUnder(folder.Folder);
        }

        private static bool IsWithinButton(DependencyObject node)
        {
            while (node != null)
            {
                if (node is Button) return true;
                node = VisualTreeHelper.GetParent(node);
            }
            return false;
        }

        /// <summary>Walk up from a clicked element to the row's <see cref="SelectableFolder"/>.</summary>
        private static SelectableFolder ResolveFolder(DependencyObject node)
        {
            while (node != null)
            {
                if (node is FrameworkElement fe && fe.DataContext is SelectableFolder sf) return sf;
                node = VisualTreeHelper.GetParent(node);
            }
            return null;
        }
    }
}
