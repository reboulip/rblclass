using System.Windows;

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
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
