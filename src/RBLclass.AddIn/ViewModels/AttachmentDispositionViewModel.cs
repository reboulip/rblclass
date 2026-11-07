using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Threading;
using RBLclass.AddIn.Localization;
using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>How the disposition modal applies choices to the attachments (v2.11.0.0 C1).</summary>
    public enum AttachmentDispositionMode
    {
        /// <summary>One action and one folder for every attachment (the default).</summary>
        SameForAll,

        /// <summary>An action per attachment, one folder for every saved one.</summary>
        SameFolder,

        /// <summary>An action and a folder per attachment.</summary>
        PerAttachment
    }

    /// <summary>
    /// View model for the F2 per-attachment disposition modal: one row per
    /// attachment across all selected mails (encrypted mails excluded and
    /// reported), each row Delete, Save-to a directory, or Keep (leave it on the
    /// filed copy, v2.5.0.0 B1). Confirm is gated until every Save-to row has a
    /// directory; Delete and Keep rows never block it.
    /// </summary>
    public sealed class AttachmentDispositionViewModel : ObservableObject
    {
        private AttachmentDispositionMode _mode = AttachmentDispositionMode.SameForAll;

        /// <param name="defaultFolder">
        /// Optional: the configured default folder, resolved off the UI thread
        /// (empty when unreachable). Pre-fills the folder when it completes,
        /// unless the user picked one meanwhile (v2.11.0.0 C1).
        /// </param>
        public AttachmentDispositionViewModel(
            IReadOnlyList<(MailItemRef Item, IReadOnlyList<AttachmentInfo> Attachments, bool IsEncrypted)> groups,
            ILocalizationService loc,
            Task<string> defaultFolder = null)
        {
            Rows = new ObservableCollection<AttachmentRowViewModel>();
            var encrypted = new List<string>();
            int inlineExcluded = 0;

            foreach (var g in groups)
            {
                if (g.IsEncrypted)
                {
                    encrypted.Add(string.IsNullOrWhiteSpace(g.Item.Subject)
                        ? loc.GetString("SentTriage_NoSubject") : g.Item.Subject);
                    continue;
                }
                foreach (var att in g.Attachments)
                {
                    // Inline/embedded images are not detachable files - never
                    // offer them for disposition, just count them (v2.5.0.0 B2).
                    if (att.IsInline) { inlineExcluded++; continue; }
                    var row = new AttachmentRowViewModel(g.Item, att, g.Item.Subject);
                    row.PropertyChanged += OnRowChanged;
                    Rows.Add(row);
                }
            }

            EncryptedNotice = encrypted.Count == 0
                ? string.Empty
                : loc.Plural(encrypted.Count,
                    "AttachDisp_EncryptedExcluded_One", "AttachDisp_EncryptedExcluded_Other");

            InlineNotice = inlineExcluded == 0
                ? string.Empty
                : loc.Plural(inlineExcluded,
                    "AttachDisp_InlineExcluded_One", "AttachDisp_InlineExcluded_Other");

            IntroText = loc.GetString("AttachDisp_Intro");
            GroupCountText = loc.Plural(Rows.Count, "AttachDisp_Group_Count_One", "AttachDisp_Group_Count_Other");

            Group = new AttachmentGroupOptionsViewModel(Rows);
            Group.Picker.PropertyChanged += OnSharedFolderChanged;
            Group.Apply();

            if (defaultFolder != null)
            {
                var dispatcher = Dispatcher.CurrentDispatcher;
                defaultFolder.ContinueWith(t =>
                {
                    if (t.Status != TaskStatus.RanToCompletion || string.IsNullOrWhiteSpace(t.Result)) return;
                    dispatcher.BeginInvoke(new Action(() => Prefill(t.Result)));
                });
            }
        }

        /// <summary>
        /// Intro text at the top of the modal. Defaults to the classify-path
        /// wording; overridden to the standalone variant for A2 (v2.8.0.0).
        /// </summary>
        public string IntroText { get; set; }

        /// <summary>
        /// Show the "remove saved attachments" checkbox — set only for the
        /// standalone A2 modal; the classify path leaves it hidden so Save-to
        /// keeps stripping as before (v2.8.0.0 A3).
        /// </summary>
        public bool ShowStripToggle { get; set; }

        /// <summary>
        /// Bound to the strip checkbox: true = a Save-to attachment is removed from
        /// the mail after saving; false = saved but left on the mail. Seeded from
        /// the Settings default (v2.8.0.0 A3).
        /// </summary>
        public bool StripSavedAttachments { get; set; } = true;

        public ObservableCollection<AttachmentRowViewModel> Rows { get; }
        public string EncryptedNotice { get; }
        public bool HasEncryptedNotice => !string.IsNullOrEmpty(EncryptedNotice);

        /// <summary>Notice that inline/embedded images were excluded (v2.5.0.0 B2).</summary>
        public string InlineNotice { get; }
        public bool HasInlineNotice => !string.IsNullOrEmpty(InlineNotice);
        public bool HasRows => Rows.Count > 0;

        /// <summary>The grouped options of "same action and folder for all"; its picker is also the shared folder.</summary>
        public AttachmentGroupOptionsViewModel Group { get; }

        /// <summary>"N attachments" header of the grouped list.</summary>
        public string GroupCountText { get; }

        /// <summary>
        /// Switching to "same for all" re-reads the group choice from the rows, and
        /// rows the user set to Delete or Keep stay as they are. Switching to "one
        /// folder" gives the shared folder to every Save-to row. Switching to "per
        /// attachment" keeps each row as it is, now editable.
        /// </summary>
        public AttachmentDispositionMode Mode
        {
            get => _mode;
            set
            {
                if (!SetProperty(ref _mode, value)) return;
                Group.IsActive = value == AttachmentDispositionMode.SameForAll;
                if (value == AttachmentDispositionMode.SameForAll) Group.Seed();
                if (value == AttachmentDispositionMode.SameFolder) StartSharedFolder();
                OnPropertyChanged(nameof(IsModeSameForAll));
                OnPropertyChanged(nameof(IsModeSameFolder));
                OnPropertyChanged(nameof(IsModePerAttachment));
                OnPropertyChanged(nameof(IsGrouped));
                OnPropertyChanged(nameof(IsPerRow));
                OnPropertyChanged(nameof(ShowSharedFolder));
                OnPropertyChanged(nameof(ShowRowFolders));
            }
        }

        public bool IsModeSameForAll
        {
            get => _mode == AttachmentDispositionMode.SameForAll;
            set { if (value) Mode = AttachmentDispositionMode.SameForAll; }
        }

        public bool IsModeSameFolder
        {
            get => _mode == AttachmentDispositionMode.SameFolder;
            set { if (value) Mode = AttachmentDispositionMode.SameFolder; }
        }

        public bool IsModePerAttachment
        {
            get => _mode == AttachmentDispositionMode.PerAttachment;
            set { if (value) Mode = AttachmentDispositionMode.PerAttachment; }
        }

        public bool IsGrouped => HasRows && _mode == AttachmentDispositionMode.SameForAll;
        public bool IsPerRow => HasRows && _mode != AttachmentDispositionMode.SameForAll;
        public bool ShowSharedFolder => HasRows && _mode == AttachmentDispositionMode.SameFolder;
        public bool ShowRowFolders => _mode == AttachmentDispositionMode.PerAttachment;

        /// <summary>True when every Save-to row has a destination directory (gates Confirm).</summary>
        public bool CanConfirm => AttachmentBulkRule.CanConfirm(Rows);

        public IReadOnlyList<AttachmentDisposition> BuildDispositions() =>
            Rows.Select(r => new AttachmentDisposition(
                r.Item, r.Info.Id, r.FileName, r.Action, r.TargetDirectory)).ToArray();

        private void Prefill(string directory)
        {
            switch (_mode)
            {
                case AttachmentDispositionMode.SameForAll:
                    Group.Prefill(directory);
                    break;
                case AttachmentDispositionMode.SameFolder:
                    if (!Group.Picker.HasTarget) Group.Picker.TargetDirectory = directory;
                    break;
                case AttachmentDispositionMode.PerAttachment:
                    foreach (var r in Rows.Where(r => r.IsSaveTo && !r.HasTarget))
                        r.TargetDirectory = directory;
                    break;
            }
        }

        private void StartSharedFolder()
        {
            if (!Group.Picker.HasTarget)
            {
                var first = Rows.FirstOrDefault(r => r.IsSaveTo && r.HasTarget);
                if (first != null) Group.Picker.TargetDirectory = first.TargetDirectory;
            }
            ApplySharedFolder();
        }

        private void ApplySharedFolder()
        {
            foreach (var r in Rows)
                r.TargetDirectory = AttachmentBulkRule.SharedFolderFor(r, Group.Picker.TargetDirectory);
        }

        private void OnSharedFolderChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SaveAsPickerViewModel.TargetDirectory)
                && _mode == AttachmentDispositionMode.SameFolder)
                ApplySharedFolder();
        }

        private void OnRowChanged(object sender, PropertyChangedEventArgs e)
        {
            // In "one folder" mode a row switched to Save to takes the shared folder.
            if (e.PropertyName == nameof(AttachmentRowViewModel.Action)
                && _mode == AttachmentDispositionMode.SameFolder
                && sender is AttachmentRowViewModel row)
                row.TargetDirectory = AttachmentBulkRule.SharedFolderFor(row, Group.Picker.TargetDirectory);

            if (e.PropertyName == nameof(AttachmentRowViewModel.Action)
                || e.PropertyName == nameof(AttachmentRowViewModel.TargetDirectory)
                || e.PropertyName == nameof(AttachmentRowViewModel.IsSaveTo))
                OnPropertyChanged(nameof(CanConfirm));
        }
    }
}
