using System;
using System.Windows.Input;
using RBLclass.AddIn.Mvvm;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// View model for the sent-item triage prompt (legacy 6c, reworked): offers
    /// Move-to-Inbox / Delete / Leave for a mail that just landed in Sent Items.
    /// Only shown when the triage setting is "Ask me each time"; a fixed mode is
    /// applied without this dialog. Picking any option raises
    /// <see cref="ChoiceMade"/>; the view closes itself. Conversation widening
    /// and the "Class" action were dropped in the rework.
    /// </summary>
    public sealed class SentItemTriageViewModel : ObservableObject
    {
        public SentItemTriageViewModel(string subject, string archiveFolderPath = null)
        {
            Subject = string.IsNullOrEmpty(subject)
                ? TaskPaneServices.Localization.GetString("SentTriage_NoSubject")
                : subject;

            HasArchiveFolder = !string.IsNullOrEmpty(archiveFolderPath);
            MoveToArchiveButtonLabel = HasArchiveFolder
                ? TaskPaneServices.Localization.GetString("SentTriage_MoveToArchive_Button", archiveFolderPath)
                : TaskPaneServices.Localization.GetString("SentTriage_MoveToArchive_Button_NoFolder");

            MoveToInboxCommand = new RelayCommand(_ => Choose(SentItemTriageAction.MoveToInbox));
            DeleteCommand = new RelayCommand(_ => Choose(SentItemTriageAction.Delete));
            LeaveCommand = new RelayCommand(_ => Choose(SentItemTriageAction.Leave));
            MoveToArchiveCommand = new RelayCommand(_ => Choose(SentItemTriageAction.MoveToArchive));
        }

        /// <summary>Subject of the mail that just landed in Sent Items.</summary>
        public string Subject { get; }

        /// <summary>The user's choice, or null until one of the buttons is clicked.</summary>
        public SentItemTriageAction? SelectedAction { get; private set; }

        /// <summary>True when an archive folder is configured, so its button is enabled (v2.7.0.0 B2).</summary>
        public bool HasArchiveFolder { get; }

        /// <summary>Label for the archive button, naming the configured folder (v2.7.0.0 B2).</summary>
        public string MoveToArchiveButtonLabel { get; }

        public ICommand MoveToInboxCommand { get; }
        public ICommand DeleteCommand { get; }
        public ICommand LeaveCommand { get; }
        public ICommand MoveToArchiveCommand { get; }

        /// <summary>Raised once a choice is made, so the hosting window can close.</summary>
        public event EventHandler ChoiceMade;

        private void Choose(SentItemTriageAction action)
        {
            SelectedAction = action;
            ChoiceMade?.Invoke(this, EventArgs.Empty);
        }
    }
}
