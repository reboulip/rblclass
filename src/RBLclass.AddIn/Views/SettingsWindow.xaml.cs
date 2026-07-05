using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using RBLclass.AddIn.Theming;
using RBLclass.AddIn.ViewModels;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// The Step 9 settings dialog (legacy §7, modernised). Every option is
    /// live-applied - <see cref="ViewModels.SettingsViewModel"/> persists each
    /// change immediately, so this window has nothing to do but close.
    /// </summary>
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            InitializeComponent();
            DialogPlacement.OwnByOutlook(this);
            ThemeService.Apply(this);
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            // Enter (IsDefault) / Esc (IsCancel) click this button without moving
            // focus off the TextBox, so its LostFocus-triggered binding never
            // fired - commit it before closing so the edit isn't lost (v2.7.1.0 E1).
            if (Keyboard.FocusedElement is TextBox tb)
                tb.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
            Close();
        }

        private void LearnBanner_Click(object sender, RoutedEventArgs e) =>
            (DataContext as SettingsViewModel)?.LearnBannerFromSelection();

        private void ClearBanner_Click(object sender, RoutedEventArgs e) =>
            (DataContext as SettingsViewModel)?.ClearBanner();

        private void DiagnoseBanner_Click(object sender, RoutedEventArgs e) =>
            (DataContext as SettingsViewModel)?.DiagnoseSelectedMail();

        private void AddFavoriteFolder_Click(object sender, RoutedEventArgs e) =>
            (DataContext as SettingsViewModel)?.AddFavoriteFolder();

        private void RemoveFavoriteFolder_Click(object sender, RoutedEventArgs e)
        {
            if (FavoriteFoldersList.SelectedItem is string path)
                (DataContext as SettingsViewModel)?.RemoveFavoriteFolder(path);
        }

        private void PickArchiveFolder_Click(object sender, RoutedEventArgs e) =>
            (DataContext as SettingsViewModel)?.PickArchiveFolder();

        private void ClearArchiveFolder_Click(object sender, RoutedEventArgs e) =>
            (DataContext as SettingsViewModel)?.ClearArchiveFolder();
    }
}
