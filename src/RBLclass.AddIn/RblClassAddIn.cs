using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using RBLclass.Core;
using RBLclass.Core.Persistence;
using RBLclass.Outlook.Adapter;
using Serilog;
using Office = Microsoft.Office.Core;
// Aliased as OutlookOM (not "Outlook") because RBLclass.Outlook.Adapter
// introduces an RBLclass.Outlook namespace that would shadow a plain alias.
using OutlookOM = Microsoft.Office.Interop.Outlook;

namespace RBLclass.AddIn
{
    /// <summary>
    /// RBLclass Outlook COM add-in shell. Thin IDTExtensibility2 +
    /// IRibbonExtensibility host; all business logic lives in RBLclass.Core and
    /// all Outlook OM access in RBLclass.Outlook.Adapter. This is the
    /// composition root: it wires the folder index (SQLite repo + Outlook
    /// adapter) and drives its lifecycle.
    /// </summary>
    [ComVisible(true)]
    [Guid(ClsidString)]
    [ProgId(ProgIdString)]
    // AutoDispatch (not None) so Office can resolve ribbon onAction callbacks
    // via IDispatch::GetIDsOfNames on the class itself (CLAUDE.md).
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class RblClassAddIn : IDTExtensibility2,
                                 Office.IRibbonExtensibility,
                                 Office.ICustomTaskPaneConsumer
    {
        // Kept in sync with release.config.json (Clsid / ProgId).
        public const string ClsidString = "43808654-547E-4222-9BE3-7FA4B781FA44";
        public const string ProgIdString = "RBLclass.AddIn";

        private OutlookOM.Application _outlookApp;

        private IFolderRepository _repository;
        private IMailStore _mailStore;
        private IFolderTree _folderTree;
        private IFolderSearch _folderSearch;
        private IClassifier _classifier;
        private ISettingsStore _settingsStore;
        private readonly ForgottenAttachmentGuard _attachmentGuard = new ForgottenAttachmentGuard();
        private readonly ExternalRecipientGuard _externalGuard = new ExternalRecipientGuard();
        private IndexResult _lastIndexResult;

        // One-shot timer that runs the first-run folder walk on the Outlook UI
        // thread shortly AFTER startup, so the COM walk stays on the STA thread
        // (CLAUDE.md) without blocking Outlook's startup.
        private Timer _firstRunWalkTimer;

        // Custom Task Pane (Office gives us the factory via ICustomTaskPaneConsumer).
        private Office.ICTPFactory _ctpFactory;
        private Office.CustomTaskPane _taskPane;

        // Kept in a field so the SelectionChange handler keeps firing (GC would
        // otherwise collect the event source - CLAUDE.md Outlook-events rule).
        private OutlookOM.Explorer _activeExplorer;

        private string _dbPath;
        private string _logDirectory;

        public void OnConnection(object application,
                                 ext_ConnectMode connectMode,
                                 object addInInst,
                                 ref Array custom)
        {
            try
            {
                _outlookApp = (OutlookOM.Application)application;
            }
            catch (Exception ex)
            {
                ShowError("OnConnection failed", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                if (_outlookApp != null)
                {
                    Marshal.FinalReleaseComObject(_outlookApp);
                    _outlookApp = null;
                }
            }
            catch (Exception ex)
            {
                ShowError("OnDisconnection failed", ex);
            }
        }

        public void OnAddInsUpdate(ref Array custom) { }

        /// <summary>
        /// Office hands us the Custom Task Pane factory here (the add-in
        /// implements ICustomTaskPaneConsumer). We keep it and create the pane
        /// lazily on first "Open folder" click.
        /// </summary>
        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            try
            {
                _ctpFactory = CTPFactoryInst;
            }
            catch (Exception ex)
            {
                TryLog("CTPFactoryAvailable failed", ex);
            }
        }

        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                InitializeStoragePaths();
                InitializeLogging();

                var version = Assembly.GetExecutingAssembly().GetName().Version;
                Log.Information(
                    "RBLclass add-in starting (version {Version}, 64-bit process={Is64}).",
                    version, Environment.Is64BitProcess);

                string connectionString = "Data Source=" + _dbPath;
                _repository = new SqliteFolderRepository(connectionString);
                _mailStore = new OutlookMailStore(_outlookApp);
                _folderTree = new FolderIndexService(_mailStore, _repository);
                _folderSearch = new FolderSearchService(_folderTree);
                _settingsStore = new SqliteSettingsStore(connectionString);
                _settingsStore.EnsureSchema();

                _classifier = new ClassifierService(_mailStore);

                // Publish services for the Custom Task Pane host control (which
                // Office instantiates via COM and so can't be constructor-injected).
                TaskPaneServices.Search = _folderSearch;
                TaskPaneServices.Settings = _settingsStore;
                TaskPaneServices.Classifier = _classifier;
                TaskPaneServices.GetSelection = () => _mailStore.GetSelectedItems();
                TaskPaneServices.CreateSubfolder = (parent, name) =>
                {
                    try
                    {
                        var created = _mailStore.CreateSubfolder(parent, name);
                        if (created != null)
                            _folderTree.ReindexStore(parent.StoreId); // targeted re-index
                        return created;
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "CreateSubfolder failed under {Path}", parent.FullPath);
                        return null;
                    }
                };
                TaskPaneServices.Navigate = (folder, newWindow) =>
                {
                    try
                    {
                        _mailStore.NavigateTo(folder.StoreId, folder.EntryId, newWindow);
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "NavigateTo failed for {Path}", folder.FullPath);
                    }
                };
                TaskPaneServices.ConfirmMarkTasksComplete = count =>
                {
                    string noun = count == 1 ? "item is" : "items are";
                    var answer = MessageBox.Show(
                        count + " selected " + noun + " flagged as an incomplete task.\n\n" +
                        "Mark them complete in the destination folder?",
                        "RBLclass - Classify",
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    switch (answer)
                    {
                        case DialogResult.Yes: return true;
                        case DialogResult.No: return false;
                        default: return null; // Cancel - abort the classify
                    }
                };
                TaskPaneServices.ConfirmSendWithoutAttachment = () =>
                {
                    var answer = MessageBox.Show(
                        "This message mentions an attachment, but none is attached.\n\n" +
                        "Send anyway?",
                        "RBLclass - Send",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    return answer == DialogResult.Yes;
                };
                TaskPaneServices.ConfirmSendToExternal = external =>
                {
                    string list = string.Join("\n", external.Select(r =>
                        string.IsNullOrEmpty(r.DisplayName) ? r.Address : r.DisplayName + " <" + r.Address + ">"));
                    string noun = external.Count == 1 ? "recipient is" : "recipients are";
                    var answer = MessageBox.Show(
                        external.Count + " " + noun + " outside the organisation:\n\n" + list +
                        "\n\nSend anyway?",
                        "RBLclass - Send",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    return answer == DialogResult.Yes;
                };

                // Keep the classify pane's selection count live.
                try
                {
                    _activeExplorer = _outlookApp.ActiveExplorer();
                    if (_activeExplorer != null)
                        _activeExplorer.SelectionChange += OnExplorerSelectionChange;
                }
                catch (Exception ex) { TryLog("SelectionChange subscribe failed", ex); }

                // Send-time guards (legacy 6a-6b). Kept on _outlookApp, which is
                // already held in a field for the add-in's lifetime (CLAUDE.md
                // Outlook-events rule - the source must not be GC'd).
                try { _outlookApp.ItemSend += Application_ItemSend; }
                catch (Exception ex) { TryLog("ItemSend subscribe failed", ex); }

                // Fast path: load the persisted index (SQLite only, no COM walk).
                _lastIndexResult = _folderTree.Load();
                Log.Information(
                    "Folder index load: {Source} ({Stores} stores, {Folders} folders).",
                    _lastIndexResult.Source,
                    _lastIndexResult.StoreCount,
                    _lastIndexResult.FolderCount);

                // First run only: schedule the live walk just after startup.
                if (_lastIndexResult.Source == IndexSource.NeedsWalk)
                    ScheduleFirstRunWalk();
            }
            catch (Exception ex)
            {
                TryLog("OnStartupComplete failed", ex);
                ShowError("OnStartupComplete failed", ex);
            }
        }

        public void OnBeginShutdown(ref Array custom)
        {
            try
            {
                if (_firstRunWalkTimer != null)
                {
                    _firstRunWalkTimer.Stop();
                    _firstRunWalkTimer.Dispose();
                    _firstRunWalkTimer = null;
                }

                if (_activeExplorer != null)
                {
                    try { _activeExplorer.SelectionChange -= OnExplorerSelectionChange; } catch { }
                    try { Marshal.ReleaseComObject(_activeExplorer); } catch { }
                    _activeExplorer = null;
                }

                if (_outlookApp != null)
                {
                    try { _outlookApp.ItemSend -= Application_ItemSend; } catch { }
                }

                if (_taskPane != null)
                {
                    try { Marshal.ReleaseComObject(_taskPane); } catch { }
                    _taskPane = null;
                }
                if (_ctpFactory != null)
                {
                    try { Marshal.ReleaseComObject(_ctpFactory); } catch { }
                    _ctpFactory = null;
                }

                Log.Information("RBLclass add-in shutting down.");
                Log.CloseAndFlush();
            }
            catch
            {
                // Never throw from shutdown.
            }
        }

        public string GetCustomUI(string ribbonId)
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                using (var stream = asm.GetManifestResourceStream("RBLclass.AddIn.Ribbon.xml"))
                {
                    if (stream == null)
                        return string.Empty;
                    using (var reader = new StreamReader(stream))
                        return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                ShowError("GetCustomUI failed", ex);
                return string.Empty;
            }
        }

        /// <summary>Ribbon "Open folder": show the pane in Open-folder mode (toggle off if already there).</summary>
        public void OnOpenFolderClick(Office.IRibbonControl control)
        {
            try { ShowPane(PaneMode.OpenFolder); }
            catch (Exception ex) { ShowError("Open folder failed", ex); }
        }

        /// <summary>Ribbon "Classify": show the pane in Classify mode (toggle off if already there).</summary>
        public void OnClassifyClick(Office.IRibbonControl control)
        {
            try { ShowPane(PaneMode.Classify); }
            catch (Exception ex) { ShowError("Classify failed", ex); }
        }

        /// <summary>
        /// Ribbon "Remove attachments": strip attachments from the current mail
        /// selection (legacy 5e standalone entry point), with confirmation.
        /// </summary>
        public void OnRemoveAttachmentsClick(Office.IRibbonControl control)
        {
            try
            {
                var items = _mailStore.GetSelectedItems();
                if (items.Count == 0)
                {
                    MessageBox.Show("Select one or more mails first.", "RBLclass",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var confirm = MessageBox.Show(
                    "Remove all attachments from " + items.Count + " mail(s)? This cannot be undone.",
                    "RBLclass - Remove attachments",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (confirm != DialogResult.Yes) return;

                int done = 0;
                foreach (var item in items)
                {
                    try { _mailStore.RemoveAttachments(item); done++; }
                    catch (Exception ex) { Log.Error(ex, "RemoveAttachments failed for an item."); }
                }

                Log.Information("Removed attachments from {Count} mail(s).", done);
                MessageBox.Show("Removed attachments from " + done + " mail(s).", "RBLclass",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError("Remove attachments failed", ex);
            }
        }

        /// <summary>Keep the classify pane's "N mails selected" label in sync with Outlook.</summary>
        private void OnExplorerSelectionChange()
        {
            try
            {
                int count = _mailStore.GetSelectedItems().Count;
                TaskPaneServices.Host?.SetSelectionCount(count);
            }
            catch (Exception ex)
            {
                TryLog("SelectionChange handler failed", ex);
            }
        }

        /// <summary>
        /// Send-time guards (legacy 6a-6b): warn on a forgotten attachment and
        /// on external recipients, each cancellable. olMail-only, like the
        /// legacy. Top-level try/catch - never let a guard failure block a send
        /// (CLAUDE.md: COM exceptions must never escape into Outlook).
        /// </summary>
        private void Application_ItemSend(object item, ref bool cancel)
        {
            try
            {
                var info = _mailStore.InspectForSend(item);
                if (info == null) return;

                var keywords = ParseList(_settingsStore.Get(
                    SettingsKeys.ForgottenAttachmentKeywords, "attach;enclos;joint;PJ"));
                if (_attachmentGuard.ShouldWarn(info.Body, info.AttachmentCount, keywords))
                {
                    if (TaskPaneServices.ConfirmSendWithoutAttachment != null &&
                        !TaskPaneServices.ConfirmSendWithoutAttachment())
                    {
                        cancel = true;
                        return;
                    }
                }

                if (_settingsStore.GetBool(SettingsKeys.SendExternalWarning, true))
                {
                    var domains = ParseList(_settingsStore.Get(SettingsKeys.InternalDomains, string.Empty));
                    var external = _externalGuard.FindExternal(info.Recipients, domains);
                    if (external.Count > 0 &&
                        TaskPaneServices.ConfirmSendToExternal != null &&
                        !TaskPaneServices.ConfirmSendToExternal(external))
                    {
                        cancel = true;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                TryLog("ItemSend guard failed", ex);
            }
        }

        /// <summary>Split a semicolon-separated settings value into trimmed, non-empty entries.</summary>
        private static IReadOnlyList<string> ParseList(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return new string[0];
            return value.Split(';')
                        .Select(s => s.Trim())
                        .Where(s => s.Length > 0)
                        .ToArray();
        }

        private void ShowPane(PaneMode mode)
        {
            EnsureTaskPane();
            if (_taskPane == null) return;

            var host = TaskPaneServices.Host;

            // Re-clicking the active mode's button hides the pane.
            if (_taskPane.Visible && host != null && host.CurrentMode == mode)
            {
                _taskPane.Visible = false;
                return;
            }

            if (host != null)
            {
                if (mode == PaneMode.Classify) host.ShowClassify();
                else host.ShowOpenFolder();
            }
            _taskPane.Visible = true;
        }

        private void EnsureTaskPane()
        {
            if (_taskPane != null) return;

            if (_ctpFactory == null)
            {
                ShowError("Task pane unavailable",
                    new InvalidOperationException(
                        "Office did not provide a Custom Task Pane factory."));
                return;
            }

            // CreateCTP instantiates the COM-registered host control by ProgId.
            _taskPane = _ctpFactory.CreateCTP(
                RblClassTaskPaneHost.ProgIdString, "RBLclass", Type.Missing);
            _taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            _taskPane.Width = 360;
        }

        /// <summary>
        /// Step 1 diagnostic: report the folder-index status so the increment
        /// can be verified on the dev machine before the search UI exists
        /// (Step 3). Replaced by the real actions in later steps.
        /// </summary>
        public void OnIndexStatusClick(Office.IRibbonControl control)
        {
            try
            {
                var r = _lastIndexResult;
                int cached = _folderTree != null ? _folderTree.GetAll().Count : 0;

                string status;
                if (r == null)
                {
                    status = "Indexing has not started yet.";
                }
                else if (r.Source == IndexSource.NeedsWalk)
                {
                    status = "First-run folder walk is scheduled / in progress." +
                             Environment.NewLine +
                             "Reopen this dialog in a moment to see the counts.";
                }
                else
                {
                    status =
                        "Index source : " + r.Source + Environment.NewLine +
                        "Stores       : " + r.StoreCount + Environment.NewLine +
                        "Folders      : " + r.FolderCount + Environment.NewLine +
                        "Cached now   : " + cached;
                }

                string message =
                    status + Environment.NewLine + Environment.NewLine +
                    "Database : " + _dbPath + Environment.NewLine +
                    "Logs     : " + _logDirectory;

                MessageBox.Show(message, "RBLclass - Folder index",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError("Index status failed", ex);
            }
        }

        // --- index lifecycle helpers ---------------------------------------

        private void ScheduleFirstRunWalk()
        {
            // Runs on the Outlook UI (STA) thread via the WinForms timer tick,
            // so COM access stays on the right thread.
            _firstRunWalkTimer = new Timer { Interval = 1500 };
            _firstRunWalkTimer.Tick += FirstRunWalkTick;
            _firstRunWalkTimer.Start();
        }

        private void FirstRunWalkTick(object sender, EventArgs e)
        {
            // One-shot.
            _firstRunWalkTimer.Stop();
            _firstRunWalkTimer.Dispose();
            _firstRunWalkTimer = null;

            try
            {
                Log.Information("No folder index found; walking live stores (first run).");
                LogStoreInventory();
                var sw = Stopwatch.StartNew();
                _lastIndexResult = _folderTree.WalkAndPersist();
                sw.Stop();
                Log.Information(
                    "First-run folder walk complete: {Stores} stores, {Folders} folders in {Ms} ms.",
                    _lastIndexResult.StoreCount,
                    _lastIndexResult.FolderCount,
                    sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "First-run folder walk failed.");
            }
        }

        /// <summary>
        /// Diagnostic: log every visible store and whether the default policy
        /// would index it, so a "0 stores walked" result can be understood
        /// (e.g. an online-mode Exchange mailbox with no data file).
        /// </summary>
        private void LogStoreInventory()
        {
            try
            {
                var stores = _mailStore.GetStores();
                Log.Information("Store inventory: {Count} store(s) visible.", stores.Count);
                foreach (var s in stores)
                {
                    Log.Information(
                        "  store: IsDataFileStore={IsData}  name='{Name}'  id={Id}",
                        s.IsDataFileStore, s.DisplayName, s.StoreId);
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Store inventory logging failed.");
            }
        }

        // --- infrastructure -------------------------------------------------

        private void InitializeStoragePaths()
        {
            string localAppData = Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData);
            string dataDir = Path.Combine(localAppData, "RBLclass");
            Directory.CreateDirectory(dataDir);

            _dbPath = Path.Combine(dataDir, "rblclass.db");
            _logDirectory = Path.Combine(dataDir, "logs");
            Directory.CreateDirectory(_logDirectory);
        }

        private void InitializeLogging()
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File(
                    Path.Combine(_logDirectory, "rblclass-.log"),
                    rollingInterval: Serilog.RollingInterval.Day,
                    retainedFileCountLimit: 14,
                    shared: true)
                .CreateLogger();
        }

        private static void TryLog(string message, Exception ex)
        {
            try { Log.Error(ex, message); } catch { /* logging not ready */ }
        }

        private static void ShowError(string title, Exception ex)
        {
            try
            {
                MessageBox.Show(ex.ToString(), "RBLclass - " + title,
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch
            {
                // Never let an exception escape into Outlook - it would disable
                // the add-in by flipping LoadBehavior 3 -> 2.
            }
        }
    }
}
