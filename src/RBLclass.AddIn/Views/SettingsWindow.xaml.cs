using System.Windows;
using RBLclass.AddIn.Theming;

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

        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
