using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using RBLclass.AddIn.Localization;
using RBLclass.AddIn.Mvvm;
using RBLclass.Core;

namespace RBLclass.AddIn.ViewModels
{
    /// <summary>
    /// View model for the F2 per-attachment disposition modal: one row per
    /// attachment across all selected mails (encrypted mails excluded and
    /// reported), each row Delete or Save-to a directory. Confirm is gated until
    /// every Save-to row has a directory.
    /// </summary>
    public sealed class AttachmentDispositionViewModel : ObservableObject
    {
        public AttachmentDispositionViewModel(
            IReadOnlyList<(MailItemRef Item, IReadOnlyList<AttachmentInfo> Attachments, bool IsEncrypted)> groups,
            ILocalizationService loc)
        {
            Rows = new ObservableCollection<AttachmentRowViewModel>();
            var encrypted = new List<string>();

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
                    var row = new AttachmentRowViewModel(g.Item, att, g.Item.Subject);
                    row.PropertyChanged += OnRowChanged;
                    Rows.Add(row);
                }
            }

            EncryptedNotice = encrypted.Count == 0
                ? string.Empty
                : loc.Plural(encrypted.Count,
                    "AttachDisp_EncryptedExcluded_One", "AttachDisp_EncryptedExcluded_Other");
        }

        public ObservableCollection<AttachmentRowViewModel> Rows { get; }
        public string EncryptedNotice { get; }
        public bool HasEncryptedNotice => !string.IsNullOrEmpty(EncryptedNotice);
        public bool HasRows => Rows.Count > 0;

        /// <summary>True when every Save-to row has a destination directory (gates Confirm).</summary>
        public bool CanConfirm =>
            Rows.All(r => r.Action != AttachmentDispositionAction.SaveTo
                          || !string.IsNullOrWhiteSpace(r.TargetDirectory));

        /// <summary>"Apply to all": set every row to Delete (the safe default).</summary>
        public void DeleteAll()
        {
            foreach (var r in Rows) r.Action = AttachmentDispositionAction.Delete;
        }

        public IReadOnlyList<AttachmentDisposition> BuildDispositions() =>
            Rows.Select(r => new AttachmentDisposition(
                r.Item, r.Info.Id, r.FileName, r.Action, r.TargetDirectory)).ToArray();

        private void OnRowChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(AttachmentRowViewModel.Action)
                || e.PropertyName == nameof(AttachmentRowViewModel.TargetDirectory)
                || e.PropertyName == nameof(AttachmentRowViewModel.IsSaveTo))
                OnPropertyChanged(nameof(CanConfirm));
        }
    }
}
