using System.Windows;
using System.Windows.Input;
using RBLclass.AddIn.Theming;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// Small modal prompt for a single line of text (used by the pane's per-row
    /// "+" to name a new sub-folder). Themed and parented to Outlook like the
    /// other dialogs. Returns the entered text via <see cref="EnteredName"/> when
    /// <see cref="Window.DialogResult"/> is true.
    /// </summary>
    public partial class NamePromptWindow : Window
    {
        public NamePromptWindow(string parentFolderName = null)
        {
            InitializeComponent();
            DialogPlacement.OwnByOutlook(this);
            ThemeService.Apply(this);

            if (!string.IsNullOrEmpty(parentFolderName))
                PromptLabel.Text = TaskPaneServices.Localization.GetString(
                    "NamePrompt_Label_WithParent", parentFolderName);

            Loaded += (s, e) => NameBox.Focus();
        }

        /// <summary>The text entered once <see cref="Window.DialogResult"/> is true.</summary>
        public string EnteredName { get; private set; }

        private void NameBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) Accept();
        }

        private void Ok_Click(object sender, RoutedEventArgs e) => Accept();

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Accept()
        {
            var name = NameBox.Text?.Trim();
            if (string.IsNullOrEmpty(name)) return;

            EnteredName = name;
            DialogResult = true;
            Close();
        }
    }
}
