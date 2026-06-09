using System;
using System.Windows.Input;
using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// View model for the sent-item triage prompt (legacy 6c): offers
    /// Class / Delete / Move-to-Inbox / Leave for a mail that just landed in
    /// Sent Items, with a "whole conversation" toggle shared with the Classify
    /// pane's setting (it's the same underlying preference). Picking any
    /// option raises <see cref="ChoiceMade"/>; the view closes itself.
    /// </summary>
    public sealed class SentItemTriageViewModel : ObservableObject
    {
        private readonly ISettingsStore _settings;
        private bool _wholeConversation;

        public SentItemTriageViewModel(string subject, ISettingsStore settings = null)
        {
            Subject = string.IsNullOrEmpty(subject) ? "(no subject)" : subject;
            _settings = settings;
            _wholeConversation = _settings?.GetBool(SettingsKeys.WidenConversation, false) ?? false;

            ClassCommand = new RelayCommand(_ => Choose(SentItemTriageAction.Class));
            DeleteCommand = new RelayCommand(_ => Choose(SentItemTriageAction.Delete));
            MoveToInboxCommand = new RelayCommand(_ => Choose(SentItemTriageAction.MoveToInbox));
            LeaveCommand = new RelayCommand(_ => Choose(SentItemTriageAction.Leave));
        }

        /// <summary>Subject of the mail that just landed in Sent Items.</summary>
        public string Subject { get; }

        public bool WholeConversation
        {
            get => _wholeConversation;
            set
            {
                if (SetProperty(ref _wholeConversation, value))
                    _settings?.SetBool(SettingsKeys.WidenConversation, value);
            }
        }

        /// <summary>The user's choice, or null until one of the buttons is clicked.</summary>
        public SentItemTriageAction? SelectedAction { get; private set; }

        public ICommand ClassCommand { get; }
        public ICommand DeleteCommand { get; }
        public ICommand MoveToInboxCommand { get; }
        public ICommand LeaveCommand { get; }

        /// <summary>Raised once a choice is made, so the hosting window can close.</summary>
        public event EventHandler ChoiceMade;

        private void Choose(SentItemTriageAction action)
        {
            SelectedAction = action;
            ChoiceMade?.Invoke(this, EventArgs.Empty);
        }
    }
}
