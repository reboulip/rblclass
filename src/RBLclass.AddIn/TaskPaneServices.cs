using System;
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
        /// <summary>Folder search service (over the cached index).</summary>
        public static IFolderSearch Search;

        /// <summary>Settings store (task-pane toggles persist here).</summary>
        public static ISettingsStore Settings;

        /// <summary>
        /// Navigate Outlook to a folder. Args: target folder, open-in-new-window.
        /// Runs on the Outlook UI thread (invoked from the pane's WPF handlers).
        /// </summary>
        public static Action<FolderNode, bool> Navigate;
    }
}
