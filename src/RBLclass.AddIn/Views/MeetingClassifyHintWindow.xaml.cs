using System.Windows;
using RBLclass.AddIn.Theming;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// One-time discoverability hint shown after a classify skipped meeting
    /// items because the "classify meeting items" setting is off (v2.9.0.0 C2).
    /// Themed and parented to Outlook like the other dialogs; the caller
    /// persists the "don't show again" choice via <see cref="DontShowAgain"/>.
    /// </summary>
    public partial class MeetingClassifyHintWindow : Window
    {
        public MeetingClassifyHintWindow()
        {
            InitializeComponent();
            DialogPlacement.OwnByOutlook(this);
            ThemeService.Apply(this);
        }

        /// <summary>True when the user ticked "don't show this again" before closing.</summary>
        public bool DontShowAgain => DontShowAgainBox.IsChecked == true;

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
