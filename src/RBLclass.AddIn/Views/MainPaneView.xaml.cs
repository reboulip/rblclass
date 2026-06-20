using System;
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

        // True when Ctrl was used as a modifier (Ctrl+A, Ctrl+C...) since it was
        // pressed - releasing it must then NOT toggle "List every matching folder".
        private bool _ctrlComboUsed;

        private void QueryBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Vm?.FileToHighlighted();
                e.Handled = true;
            }
        }

        private void QueryBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            var key = e.Key == Key.System ? e.SystemKey : e.Key;
            if (key == Key.LeftCtrl || key == Key.RightCtrl)
            {
                if (!e.IsRepeat) _ctrlComboUsed = false; // fresh Ctrl press
            }
            else if ((Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                _ctrlComboUsed = true; // Ctrl is acting as a modifier
            }
            else if (key == Key.Tab && (Keyboard.Modifiers & ModifierKeys.Shift) == 0)
            {
                // Tab from the query box opens the Options panel and moves
                // focus to its first checkbox (v2.4 B1). The focus is deferred:
                // the checkbox is in a just-made-visible panel and is not yet
                // realized in this frame, so Focus() here would no-op.
                if (Vm != null) Vm.IsOptionsExpanded = true;
                Dispatcher.BeginInvoke(
                    System.Windows.Threading.DispatcherPriority.Input,
                    new Action(() => FirstOptionCheckBox?.Focus()));
                e.Handled = true;
            }
        }

        private void Options_Click(object sender, RoutedEventArgs e)
        {
            if (Vm != null) Vm.IsOptionsExpanded = !Vm.IsOptionsExpanded;
        }

        /// <summary>Ctrl pressed and released on its own toggles "List every matching folder" (v2.2).</summary>
        private void QueryBox_KeyUp(object sender, KeyEventArgs e)
        {
            if ((e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
                && (Keyboard.Modifiers & ModifierKeys.Control) == 0)
            {
                if (!_ctrlComboUsed && Vm != null)
                    Vm.AllResults = !Vm.AllResults;
                _ctrlComboUsed = false;
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

        private void RemoveDestination_Click(object sender, RoutedEventArgs e)
        {
            var folder = (sender as FrameworkElement)?.DataContext as RBLclass.Core.FolderNode;
            if (folder != null) Vm?.RemoveDestination(folder);
        }

        private void ClearSelection_Click(object sender, RoutedEventArgs e) => Vm?.ClearDestinations();

        private void Undo_Click(object sender, RoutedEventArgs e) => Vm?.UndoLast();

        private void AutoClass_Click(object sender, RoutedEventArgs e) => Vm?.AutoClassSelected();

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
