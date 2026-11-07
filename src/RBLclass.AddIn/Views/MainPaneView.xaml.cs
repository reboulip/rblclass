using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using RBLclass.AddIn.ViewModels;
using Serilog;

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
            SizeChanged += MainPane_SizeChanged;
        }

        private MainPaneViewModel Vm => DataContext as MainPaneViewModel;

        /// <summary>
        /// Switches the action-button bar between the stacked (portrait) and
        /// single-row (landscape) layouts based on the pane's own aspect ratio,
        /// not a user setting (v2.9.0.0 A2).
        /// </summary>
        private void MainPane_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (Vm != null) Vm.IsHorizontalLayout = e.NewSize.Width > e.NewSize.Height;
        }

        /// <summary>
        /// Put keyboard focus in the search box (v2.7.0.0 A2), called whenever the
        /// pane is shown so the user can type a query straight away. Deferred to
        /// Input priority because on a fresh show the TextBox may not be realized
        /// yet, and selects existing text so a new query overwrites it.
        /// </summary>
        public void FocusSearch()
        {
            Dispatcher.BeginInvoke(
                new Action(() => FocusSearchCore(requireForeground: false)),
                System.Windows.Threading.DispatcherPriority.Input);
        }

        private void FocusSearchCore(bool requireForeground)
        {
            try
            {
                if (ComponentDispatcher.IsThreadModal || !QueryBox.IsEnabled) return;

                var source = PresentationSource.FromVisual(QueryBox) as HwndSource;
                bool foreground = source != null && PaneFocus.IsRootForeground(source.Handle);
                if (requireForeground && !foreground) return;

                MoveWpfFocusToSearch();
                if (foreground && !PaneFocus.HasFocus(source.Handle))
                {
                    Log.Debug("Search focus: WPF focus did not take, moving Win32 focus to the pane.");
                    PaneFocus.SetFocus(source.Handle);
                    MoveWpfFocusToSearch();
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Moving focus to the search box failed.");
            }
        }

        private void MoveWpfFocusToSearch()
        {
            QueryBox.Focus();
            Keyboard.Focus(QueryBox);
            QueryBox.SelectAll();
        }

        // True when Ctrl was used as a modifier (Ctrl+A, Ctrl+C...) since it was
        // pressed - releasing it must then NOT toggle "List every matching folder".
        private bool _ctrlComboUsed;

        private async void QueryBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                // Ctrl+Enter opens the highlighted folder (v2.7.0.0 A2), distinct
                // from plain Enter which files into it. Deferred to ApplicationIdle
                // (B2, v2.9.0.0): called synchronously here, the TextBox still holds
                // Win32 keyboard focus and Ctrl is still physically down mid-keystroke,
                // which was silently swallowing the downstream Explorer.Activate() COM
                // call and leaving Delete dead in the message list. The per-row open
                // button doesn't hit this - by the time a Button.Click handler runs,
                // the click gesture has already fully settled.
                e.Handled = true;
                _ = Dispatcher.BeginInvoke(
                    new Action(() =>
                    {
                        try { Vm?.OpenHighlighted(); }
                        catch (Exception ex) { Log.Error(ex, "OpenHighlighted failed."); }
                        FocusMessageListWhenIdle();
                    }),
                    System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                return;
            }

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                try { await (Vm?.FileToHighlighted() ?? Task.CompletedTask); }
                catch (Exception ex) { Log.Error(ex, "FileToHighlighted failed."); }
                ReturnFocusToSearch();
            }
        }

        /// <summary>
        /// Return keyboard focus to the pane's search box after a pane action or a
        /// modal close (v2.7.0.0 A1, v2.11.0.0 B3), so the user can type or press
        /// Ctrl+Z straight away - Outlook grabs focus for the message list when
        /// the filed item leaves the view. Deferred to ApplicationIdle, never
        /// Input: it must run after Outlook's post-move selection change and after
        /// IsBusy clears at Background priority, since the box is disabled while
        /// busy. Skipped unless the pane's window is in the foreground, so it
        /// cannot steal focus from a compose window.
        /// "Open folder" deliberately does not call this: focus stays in the
        /// explorer so the Delete key keeps working (#49).
        /// </summary>
        public void ReturnFocusToSearch()
        {
            Dispatcher.BeginInvoke(
                new Action(() => FocusSearchCore(requireForeground: true)),
                System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        }

        /// <summary>
        /// Hand the keyboard to Outlook's message list after opening a folder, so
        /// Delete acts on the selected mail (#49). Skipped when the folder opened
        /// in a new window, which leaves the pane's window in the background.
        /// </summary>
        private void FocusMessageListWhenIdle()
        {
            Dispatcher.BeginInvoke(
                new Action(() =>
                {
                    try
                    {
                        var source = PresentationSource.FromVisual(QueryBox) as HwndSource;
                        if (source == null || !PaneFocus.IsRootForeground(source.Handle)) return;
                        if (!PaneFocus.FocusMessageList(source.Handle))
                            Log.Debug("Message list window not found; focus left as it is.");
                    }
                    catch (Exception ex)
                    {
                        Log.Warning(ex, "Moving focus to the message list failed.");
                    }
                }),
                System.Windows.Threading.DispatcherPriority.ApplicationIdle);
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

        /// <summary>
        /// Ctrl+Z anywhere in the pane fires the pane's Undo action — including
        /// while the query TextBox has focus (the classify undo is what the user
        /// wants there, not text-edit undo). Uses the tunnel (Preview) phase so
        /// it is seen before any focused child handles the chord; Ctrl+Shift+Z
        /// (redo) is deliberately not intercepted.
        /// </summary>
        private void MainPane_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Z && (Keyboard.Modifiers & ModifierKeys.Control) != 0
                && (Keyboard.Modifiers & ModifierKeys.Shift) == 0)
            {
                var vm = Vm;
                if (vm == null || !vm.CanUndo) return;

                e.Handled = true;

                // Handled here, so QueryBox_PreviewKeyDown never sees the chord:
                // without this, releasing Ctrl in the refocused box would count
                // as a lone tap and flip "List every matching folder".
                _ctrlComboUsed = true;

                try { vm.UndoLast(); }
                catch (Exception ex) { Log.Error(ex, "UndoLast failed."); }
                ReturnFocusToSearch();
            }
        }

        /// <summary>
        /// "/" jumps to the search box from anywhere in the pane and is swallowed;
        /// inside the box it types normally. Uses text input, not KeyDown, so it
        /// works on AZERTY and the numeric keypad. Pane-scoped: there is no global
        /// hook (EDR risk).
        /// </summary>
        private void MainPane_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (e.Text != "/" || QueryBox.IsKeyboardFocused) return;

            e.Handled = true;
            FocusSearchCore(requireForeground: false);
        }

        private void Options_Click(object sender, RoutedEventArgs e)
        {
            if (Vm != null) Vm.IsOptionsExpanded = !Vm.IsOptionsExpanded;
        }

        private void Journal_Click(object sender, RoutedEventArgs e)
        {
            if (Vm != null) Vm.IsStatusLogExpanded = !Vm.IsStatusLogExpanded;
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

        private async void ResultsList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Ignore double-clicks that land on the per-row buttons.
            if (IsWithinButton(e.OriginalSource as DependencyObject)) return;

            // File into the row that was actually double-clicked, not the list's
            // current selection (clicking a row toggles its checkbox without
            // changing the ListBox selection).
            var folder = ResolveFolder(e.OriginalSource as DependencyObject);
            if (folder == null) return;
            try { await (Vm?.FileToFolder(folder.Folder) ?? Task.CompletedTask); }
            catch (Exception ex) { Log.Error(ex, "FileToFolder failed."); }
            ReturnFocusToSearch();
        }

        private async void Classify_Click(object sender, RoutedEventArgs e)
        {
            try { await (Vm?.ClassifyChecked() ?? Task.CompletedTask); }
            catch (Exception ex) { Log.Error(ex, "ClassifyChecked failed."); }
            ReturnFocusToSearch();
        }

        private void RemoveDestination_Click(object sender, RoutedEventArgs e)
        {
            var folder = (sender as FrameworkElement)?.DataContext as RBLclass.Core.FolderNode;
            if (folder != null) Vm?.RemoveDestination(folder);
            ReturnFocusToSearch();
        }

        private void ClearSelection_Click(object sender, RoutedEventArgs e)
        {
            Vm?.ClearDestinations();
            ReturnFocusToSearch();
        }

        private void Undo_Click(object sender, RoutedEventArgs e)
        {
            try { Vm?.UndoLast(); }
            catch (Exception ex) { Log.Error(ex, "UndoLast failed."); }
            ReturnFocusToSearch();
        }

        private void DismissUpdate_Click(object sender, RoutedEventArgs e)
        {
            Vm?.DismissUpdateBanner();
            ReturnFocusToSearch();
        }

        private void DeletePinned_Click(object sender, RoutedEventArgs e)
        {
            try { Vm?.DeletePinned(); }
            catch (Exception ex) { Log.Error(ex, "DeletePinned failed."); }
            ReturnFocusToSearch();
        }

        private void AutoClass_Click(object sender, RoutedEventArgs e)
        {
            try { Vm?.AutoClassSelected(); }
            catch (Exception ex) { Log.Error(ex, "AutoClassSelected failed."); }
            ReturnFocusToSearch();
        }

        private void SaveAs_Click(object sender, RoutedEventArgs e)
        {
            try { Vm?.SaveAsSelected(); }
            catch (Exception ex) { Log.Error(ex, "SaveAsSelected failed."); }
            ReturnFocusToSearch();
        }

        private void ManageAttachments_Click(object sender, RoutedEventArgs e)
        {
            try { Vm?.ManageAttachmentsSelected(); }
            catch (Exception ex) { Log.Error(ex, "ManageAttachmentsSelected failed."); }
            ReturnFocusToSearch();
        }

        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            try { Vm?.OpenSettings(); }
            catch (Exception ex) { Log.Error(ex, "Settings_Click failed."); }
            ReturnFocusToSearch();
        }

        private void Help_Click(object sender, RoutedEventArgs e)
        {
            try { Vm?.OpenDocumentation(); }
            catch (Exception ex) { Log.Error(ex, "Help_Click failed."); }
            ReturnFocusToSearch();
        }

        private void UpdateFolders_Click(object sender, RoutedEventArgs e)
        {
            try { Vm?.RefreshFolders(); }
            catch (Exception ex) { Log.Error(ex, "UpdateFolders_Click failed."); }
            ReturnFocusToSearch();
        }

        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            var folder = (sender as FrameworkElement)?.DataContext as SelectableFolder;
            if (folder != null) Vm?.OpenFolder(folder.Folder);
            FocusMessageListWhenIdle();
        }

        private void NewSubfolder_Click(object sender, RoutedEventArgs e)
        {
            var folder = (sender as FrameworkElement)?.DataContext as SelectableFolder;
            if (folder != null) Vm?.CreateSubfolderUnder(folder.Folder);
            ReturnFocusToSearch();
        }

        private void RenameFolder_Click(object sender, RoutedEventArgs e)
        {
            var folder = (sender as FrameworkElement)?.DataContext as SelectableFolder;
            if (folder != null) Vm?.RenameFolder(folder.Folder);
            ReturnFocusToSearch();
        }

        private void DeleteFolder_Click(object sender, RoutedEventArgs e)
        {
            var folder = (sender as FrameworkElement)?.DataContext as SelectableFolder;
            if (folder != null) Vm?.DeleteFolder(folder.Folder);
            ReturnFocusToSearch();
        }

        private void ExpandPath_Click(object sender, RoutedEventArgs e)
        {
            var folder = (sender as FrameworkElement)?.DataContext as SelectableFolder;
            folder?.CycleDisplayMode();
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
