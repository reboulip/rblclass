using System;
using System.Collections.Generic;
using RBLclass.AddIn.Localization;
using RBLclass.Core;

namespace RBLclass.AddIn
{
    /// <summary>
    /// Composition hand-off for the Custom Task Pane. The pane host control is
    /// instantiated by Office via COM (parameterless ctor), so it cannot receive
    /// services through its constructor; the add-in's composition root publishes
    /// them here at startup and the control reads them when created. Single
    /// add-in per process, so a static holder is sufficient.
    /// </summary>
    internal static class TaskPaneServices
    {
        /// <summary>
        /// UI string lookup for the language resolved once at startup. Set
        /// before any view or window is constructed, so XAML's
        /// <see cref="LocExtension"/> and ViewModel constructors can rely on
        /// it being non-null.
        /// </summary>
        public static ILocalizationService Localization;

        /// <summary>Folder search service (over the cached index).</summary>
        public static IFolderSearch Search;

        /// <summary>Settings store (task-pane toggles persist here).</summary>
        public static ISettingsStore Settings;

        /// <summary>Classify (file) selected mail into folders.</summary>
        public static IClassifier Classifier;

        /// <summary>
        /// Snapshot of the live folder index, so the pane can resolve Auto-class
        /// remembered destinations against current folders (and detect stale
        /// ones). Read on the Outlook UI thread.
        /// </summary>
        public static Func<IReadOnlyList<FolderNode>> GetAllFolders;

        /// <summary>
        /// The folder index service, published at startup so the pane can
        /// subscribe to <see cref="IFolderIndexService.IndexStatus"/> changes and
        /// drive the header's colored status dot.
        /// </summary>
        public static IFolderIndexService FolderIndex;

        /// <summary>Read the current Outlook explorer mail selection (UI thread).</summary>
        public static Func<IReadOnlyList<MailItemRef>> GetSelection;

        /// <summary>
        /// True while the pane is running a move batch (classify / auto-class /
        /// undo). During a batch the inter-item pump-yield (D1/D2) lets Outlook
        /// repaint and fire <c>SelectionChange</c> after every <c>Move</c>; the
        /// add-in's selection handler honours this flag to skip its per-item
        /// <c>GetSelectedItems</c> COM walk, which is redundant until the batch
        /// ends (the pane does one deferred refresh then). All access is on the
        /// single STA thread. Does NOT change the yield itself - Stormshield's
        /// stable scan window after each Move is preserved.
        /// </summary>
        public static bool BatchInProgress;

        /// <summary>
        /// Create a sub-folder under a parent and re-index that store; returns
        /// the new node (or null). Args: parent folder, new name. UI thread.
        /// </summary>
        public static Func<FolderNode, string, FolderNode> CreateSubfolder;

        /// <summary>
        /// Navigate Outlook to a folder. Args: target folder, open-in-new-window.
        /// Runs on the Outlook UI thread (invoked from the pane's WPF handlers).
        /// </summary>
        public static Action<FolderNode, bool> Navigate;

        /// <summary>
        /// Ask the user whether to mark flagged-incomplete items complete in
        /// the destination before filing (legacy 5b step 3 task-completion
        /// guard). Arg: how many items qualify. Returns true = mark them
        /// complete, false = leave them as-is, null = cancel the classify
        /// entirely. A WinForms MessageBox (Yes/No/Cancel) under the hood -
        /// kept out of the WPF-agnostic view model as a plain delegate.
        /// </summary>
        public static Func<int, bool?> ConfirmMarkTasksComplete;

        /// <summary>
        /// Ask the user whether to send anyway when the body mentions an
        /// attachment but none is attached (legacy 6b forgotten-attachment
        /// guard). Returns true = send, false = cancel and go back to attach.
        /// </summary>
        public static Func<bool> ConfirmSendWithoutAttachment;

        /// <summary>
        /// Ask the user whether to send anyway to the listed external
        /// recipients (legacy 6a external-recipient guard). Returns
        /// true = send, false = cancel.
        /// </summary>
        public static Func<IReadOnlyList<RecipientAddress>, bool> ConfirmSendToExternal;

        /// <summary>
        /// Prompt the user for a single line of text (the new sub-folder name),
        /// pre-filled with the parent folder name as context. Returns null/empty
        /// when cancelled. A small modal dialog under the hood - kept out of the
        /// WPF-agnostic view model as a plain delegate.
        /// </summary>
        public static Func<string, string> PromptForName;

        /// <summary>
        /// The live pane host control (set by its ctor), so ribbon callbacks can
        /// show/refresh the single task pane.
        /// </summary>
        public static RblClassTaskPaneHost Host;

        /// <summary>
        /// Pin a just-sent mail (moved to the Inbox by sent-item triage) as the
        /// next classify target in the main pane (v2.4.0.0 E1). Wired to the
        /// pane view model when the pane is first created, so it is null until
        /// then. The pin clears after the next classify or on selection change.
        /// </summary>
        public static Action<MailItemRef> PinMailForClassify;

        /// <summary>
        /// The favourite-folder filesystem index (v2.4.0.0 F1): keyword search
        /// over the user's pre-indexed save-to directories.
        /// </summary>
        public static FavoriteFolderService FavoriteFolderService;

        /// <summary>
        /// Show the OS folder-browse dialog and return the selected path, or null
        /// when cancelled (v2.4.0.0 F1). WinForms FolderBrowserDialog under the
        /// hood; call on the UI thread.
        /// </summary>
        public static Func<string> BrowseForFolder;

        /// <summary>
        /// Activate the pig easter egg in the main pane's Auto-Class button label.
        /// Wired at startup; null until then.
        /// </summary>
        public static Action ActivateEasterEgg;

        /// <summary>
        /// Gather each item's attachments (and whether it is encrypted) for the
        /// F2 disposition modal. Touches COM - invoked on the UI thread.
        /// </summary>
        public static Func<IReadOnlyList<MailItemRef>,
            IReadOnlyList<(MailItemRef Item, IReadOnlyList<AttachmentInfo> Attachments, bool IsEncrypted)>>
            GatherAttachments;

        /// <summary>
        /// Show the F2 per-attachment disposition modal for the gathered groups
        /// and return the user's choices, or null when cancelled (which aborts
        /// the classify).
        /// </summary>
        public static Func<
            IReadOnlyList<(MailItemRef Item, IReadOnlyList<AttachmentInfo> Attachments, bool IsEncrypted)>,
            IReadOnlyList<AttachmentDisposition>> ShowAttachmentDisposition;
    }
}
