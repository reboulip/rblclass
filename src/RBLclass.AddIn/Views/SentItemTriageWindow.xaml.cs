using System;
using RBLclass.AddIn.Theming;
using RBLclass.AddIn.ViewModels;

namespace RBLclass.AddIn.Views
{
    /// <summary>
    /// The sent-item triage prompt (legacy 6c). A small modal window - the
    /// first of its kind in this add-in (every other prompt so far is a
    /// WinForms MessageBox; this one needs four custom actions plus a toggle,
    /// which MessageBoxButtons cannot express). Closes itself once the view
    /// model reports a choice.
    /// </summary>
    public partial class SentItemTriageWindow : System.Windows.Window
    {
        public SentItemTriageWindow()
        {
            InitializeComponent();
            DialogPlacement.OwnByOutlook(this);
            ThemeService.Apply(this);
            DataContextChanged += OnDataContextChanged;
        }

        private void OnDataContextChanged(object sender, System.Windows.DependencyPropertyChangedEventArgs e)
        {
            if (e.OldValue is SentItemTriageViewModel oldVm) oldVm.ChoiceMade -= OnChoiceMade;
            if (e.NewValue is SentItemTriageViewModel newVm) newVm.ChoiceMade += OnChoiceMade;
        }

        private void OnChoiceMade(object sender, EventArgs e) => Close();
    }
}
