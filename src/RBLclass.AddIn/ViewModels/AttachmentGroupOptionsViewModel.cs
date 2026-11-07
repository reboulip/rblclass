using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// The grouped view of the attachment disposition modal (v2.11.0.0 C1): one
    /// action and one folder, applied to every row that is not pinned
    /// (<see cref="AttachmentBulkRule"/>). The folder picker is the same
    /// favourite-search and browse UX as "Save as".
    /// </summary>
    public sealed class AttachmentGroupOptionsViewModel : ObservableObject
    {
        private readonly IReadOnlyList<AttachmentRowViewModel> _rows;
        private AttachmentDispositionAction _action = AttachmentDispositionAction.SaveTo;
        private bool _settingDirectory;

        public AttachmentGroupOptionsViewModel(IReadOnlyList<AttachmentRowViewModel> rows)
        {
            _rows = rows;
            Picker = new SaveAsPickerViewModel();
            Picker.PropertyChanged += OnPickerChanged;
        }

        public SaveAsPickerViewModel Picker { get; }

        /// <summary>
        /// True in "same action and folder for all" mode. Otherwise the picker is
        /// still the shared folder, but no group action is applied to the rows.
        /// </summary>
        public bool IsActive { get; set; } = true;

        public bool IsDelete
        {
            get => _action == AttachmentDispositionAction.Delete;
            set { if (value) Choose(AttachmentDispositionAction.Delete); }
        }

        public bool IsSaveTo
        {
            get => _action == AttachmentDispositionAction.SaveTo;
            set { if (value) Choose(AttachmentDispositionAction.SaveTo); }
        }

        public bool IsKeep
        {
            get => _action == AttachmentDispositionAction.Keep;
            set { if (value) Choose(AttachmentDispositionAction.Keep); }
        }

        /// <summary>Save to is chosen but no folder yet: Confirm stays disabled.</summary>
        public bool ShowNoFolder => IsSaveTo && !Picker.HasTarget;

        /// <summary>Re-read the choice from the rows, then apply it (switching back to the grouped view).</summary>
        public void Seed()
        {
            var action = AttachmentBulkRule.SeedGroupChoice(_rows, out string directory);
            SetDirectory(directory);
            SetAction(action);
            Apply();
        }

        /// <summary>
        /// Pre-fill the configured default folder when the user has not chosen
        /// one meanwhile. Does not change the chosen action.
        /// </summary>
        public void Prefill(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory) || Picker.HasTarget) return;
            SetDirectory(directory);
            Apply();
        }

        public void Apply()
        {
            if (!IsActive) return;
            foreach (var row in _rows.Where(AttachmentBulkRule.AppliesTo))
                row.ApplyBulk(_action, Picker.TargetDirectory, markUserSet: false);
        }

        private void Choose(AttachmentDispositionAction action)
        {
            SetAction(action);
            Apply();
        }

        private void SetAction(AttachmentDispositionAction action)
        {
            if (_action == action) return;
            _action = action;
            OnPropertyChanged(nameof(IsDelete));
            OnPropertyChanged(nameof(IsSaveTo));
            OnPropertyChanged(nameof(IsKeep));
            OnPropertyChanged(nameof(ShowNoFolder));
        }

        private void SetDirectory(string directory)
        {
            _settingDirectory = true;
            try { Picker.TargetDirectory = directory; }
            finally { _settingDirectory = false; }
        }

        // A folder picked by the user (browse or favourite search) means Save to.
        private void OnPickerChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(SaveAsPickerViewModel.TargetDirectory)) return;
            OnPropertyChanged(nameof(ShowNoFolder));
            if (_settingDirectory || !IsActive) return;
            if (Picker.HasTarget) SetAction(AttachmentDispositionAction.SaveTo);
            Apply();
        }
    }
}
