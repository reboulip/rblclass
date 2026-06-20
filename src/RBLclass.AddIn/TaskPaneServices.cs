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
    }
}
