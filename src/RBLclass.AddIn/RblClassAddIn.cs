using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Threading;
using Extensibility;
using RBLclass.AddIn.Localization;
using RBLclass.AddIn.ViewModels;
using RBLclass.AddIn.Views;
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
        private IFolderIndexService _folderTree;
        private IFolderSearch _folderSearch;
        private IClassifier _classifier;
        private IClassificationHistory _classificationHistory;
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

        // Sent-item triage (legacy 6c): kept alive for the same reason as
        // _activeExplorer - both the folder and its Items collection, so the
        // ItemAdd subscription keeps firing.
        private OutlookOM.Folder _sentItemsFolder;
        private OutlookOM.Items _sentItems;

        // Re-entrancy guard for the handler above (legacy "Set colSentItems =
        // Nothing" detach/reattach dance, replaced per the roadmap with a
        // plain flag): classifying or moving the triaged item makes a copy
        // transit Sent Items, which would otherwise re-trigger the prompt on
        // that transient copy.
        private bool _suppressSentItemTriage;

        // New-inspector subscription for the v2.2 reply/forward banner strip.
        // Held in a field for the add-in's lifetime (GC rule - CLAUDE.md).
        private OutlookOM.Inspectors _inspectors;

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

                // Resolve the UI language once, before any pane/window is
                // created - LocExtension and ViewModel constructors read
                // TaskPaneServices.Localization (CLAUDE.md / i18n plan).
                var preferredLanguage = Settings.Load(_settingsStore).PreferredUiLanguage;
                var outlookLanguage = new OutlookUiLanguageProvider(_outlookApp).GetOutlookUiLanguageCode();
                var resolvedLanguage = UiLanguageResolver.Resolve(preferredLanguage, outlookLanguage);
                TaskPaneServices.Localization = new LocalizationService(resolvedLanguage);

                _classificationHistory = new SqliteClassificationHistory(connectionString);
                _classificationHistory.EnsureSchema();

                _classifier = new ClassifierService(_mailStore, _classificationHistory);

                // Publish services for the Custom Task Pane host control (which
                // Office instantiates via COM and so can't be constructor-injected).
                TaskPaneServices.Search = _folderSearch;
                TaskPaneServices.Settings = _settingsStore;
                TaskPaneServices.Classifier = _classifier;
                TaskPaneServices.GetSelection = () => _mailStore.GetSelectedItems();
                TaskPaneServices.GetAllFolders = () => _folderTree.GetAll();
                TaskPaneServices.FolderIndex = _folderTree;
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
                TaskPaneServices.PromptForName = parentFolderName =>
                {
                    var prompt = new NamePromptWindow(parentFolderName);
                    return prompt.ShowDialog() == true ? prompt.EnteredName : null;
                };
                TaskPaneServices.ConfirmMarkTasksComplete = count =>
                {
                    var loc = TaskPaneServices.Localization;
                    string body = loc.Plural(count,
                        "MsgBox_ConfirmMarkTasksComplete_Body_One",
                        "MsgBox_ConfirmMarkTasksComplete_Body_Other");
                    var answer = MessageBox.Show(
                        body,
                        loc.GetString("MsgBox_ConfirmMarkTasksComplete_Title"),
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
                    var loc = TaskPaneServices.Localization;
                    var answer = MessageBox.Show(
                        loc.GetString("MsgBox_ConfirmSendWithoutAttachment_Body"),
                        loc.GetString("MsgBox_ConfirmSendWithoutAttachment_Title"),
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    return answer == DialogResult.Yes;
                };
                TaskPaneServices.ConfirmSendToExternal = external =>
                {
                    var loc = TaskPaneServices.Localization;
                    string list = string.Join("\n", external.Select(r =>
                        string.IsNullOrEmpty(r.DisplayName) ? r.Address : r.DisplayName + " <" + r.Address + ">"));
                    string body = loc.Plural(external.Count,
                        "MsgBox_ConfirmSendToExternal_Body_One",
                        "MsgBox_ConfirmSendToExternal_Body_Other",
                        list);
                    var answer = MessageBox.Show(
                        body,
                        loc.GetString("MsgBox_ConfirmSendToExternal_Title"),
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

                // Sent-item triage (legacy 6c). Folder + Items kept in fields
                // for the add-in's lifetime, mirroring _activeExplorer - GC
                // would otherwise silently unsubscribe (CLAUDE.md).
                try
                {
                    using (var session = new ComRef<OutlookOM.NameSpace>(_outlookApp.Session))
                        _sentItemsFolder = (OutlookOM.Folder)session.Value.GetDefaultFolder(
                            OutlookOM.OlDefaultFolders.olFolderSentMail);

                    if (_sentItemsFolder != null)
                    {
                        _sentItems = _sentItemsFolder.Items;
                        _sentItems.ItemAdd += SentItems_ItemAdd;
                    }
                }
                catch (Exception ex) { TryLog("Sent Items ItemAdd subscribe failed", ex); }

                // Reply/forward banner strip (v2.2). Inspectors held in a field
                // for the add-in's lifetime, like the other event sources.
                try
                {
                    _inspectors = _outlookApp.Inspectors;
                    _inspectors.NewInspector += OnNewInspector;
                }
                catch (Exception ex) { TryLog("NewInspector subscribe failed", ex); }

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

                if (_inspectors != null)
                {
                    try { _inspectors.NewInspector -= OnNewInspector; } catch { }
                    try { Marshal.ReleaseComObject(_inspectors); } catch { }
                    _inspectors = null;
                }

                if (_sentItems != null)
                {
                    try { _sentItems.ItemAdd -= SentItems_ItemAdd; } catch { }
                    try { Marshal.ReleaseComObject(_sentItems); } catch { }
                    _sentItems = null;
                }
                if (_sentItemsFolder != null)
                {
                    try { Marshal.ReleaseComObject(_sentItemsFolder); } catch { }
                    _sentItemsFolder = null;
                }

                if (_taskPane != null)
                {
                    try { _taskPane.DockPositionStateChange -= OnTaskPaneDockPositionStateChange; } catch { }
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
                string resourceName;
                switch (TaskPaneServices.Localization?.CurrentLanguage)
                {
                    case "fr": resourceName = "RBLclass.AddIn.Ribbon.fr.xml"; break;
                    case "de": resourceName = "RBLclass.AddIn.Ribbon.de.xml"; break;
                    default: resourceName = "RBLclass.AddIn.Ribbon.xml"; break;
                }

                using (var stream = asm.GetManifestResourceStream(resourceName))
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

        /// <summary>Ribbon "RBLclass pane": show or hide the single unified task pane.</summary>
        public void OnTogglePaneClick(Office.IRibbonControl control)
        {
            try { TogglePane(); }
            catch (Exception ex) { ShowError("RBLclass pane failed", ex); }
        }

        /// <summary>
        /// Ribbon "Remove attachments": strip attachments from the current mail
        /// selection (legacy 5e standalone entry point), with confirmation.
        /// </summary>
        public void OnRemoveAttachmentsClick(Office.IRibbonControl control)
        {
            try
            {
                var loc = TaskPaneServices.Localization;
                var items = _mailStore.GetSelectedItems();
                if (items.Count == 0)
                {
                    MessageBox.Show(loc.GetString("MsgBox_RemoveAttachments_SelectFirst"),
                                    loc.GetString("MsgBox_Info_Title"),
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var confirm = MessageBox.Show(
                    loc.GetString("MsgBox_RemoveAttachments_Confirm", items.Count),
                    loc.GetString("MsgBox_RemoveAttachments_Title"),
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (confirm != DialogResult.Yes) return;

                int done = 0, skippedEncrypted = 0;
                foreach (var item in items)
                {
                    try
                    {
                        if (_mailStore.RemoveAttachments(item)) done++;
                        else skippedEncrypted++; // encrypted/signed - never stripped
                    }
                    catch (Exception ex) { Log.Error(ex, "RemoveAttachments failed for an item."); }
                }

                Log.Information("Removed attachments from {Count} mail(s) ({Skipped} encrypted skipped).",
                                done, skippedEncrypted);
                string summary = loc.GetString("MsgBox_RemoveAttachments_Summary", done);
                if (skippedEncrypted > 0)
                    summary += loc.Plural(skippedEncrypted,
                        "MsgBox_RemoveAttachments_SkippedEncrypted_One",
                        "MsgBox_RemoveAttachments_SkippedEncrypted_Other");
                MessageBox.Show(summary, loc.GetString("MsgBox_Info_Title"),
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError("Remove attachments failed", ex);
            }
        }

        /// <summary>
        /// Reply/forward banner strip (v2.2): when the toggle is on and a banner
        /// has been learned, strip it from the new draft so it isn't quoted back.
        /// The draft (reply/reply-all/forward) quotes the original including its
        /// banner; stripping is a no-op for a plain new mail. Best-effort and
        /// never throws into Outlook.
        /// </summary>
        private void OnNewInspector(OutlookOM.Inspector inspector)
        {
            object item = null;
            try
            {
                var settings = Settings.Load(_settingsStore);
                if (!settings.StripBannerOnReply ||
                    string.IsNullOrWhiteSpace(settings.ExternalBannerSignature))
                    return;

                try { item = inspector.CurrentItem; } catch { item = null; }
                if (item == null) return;

                _mailStore.StripBannerFromDraft(item, settings.ExternalBannerSignature);
            }
            catch (Exception ex)
            {
                TryLog("NewInspector banner strip failed", ex);
            }
            finally
            {
                if (item != null)
                {
                    try { Marshal.ReleaseComObject(item); } catch { }
                }
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

        /// <summary>
        /// Sent-item triage (legacy 6c, reworked): on a fresh item in Sent Items,
        /// apply the configured <see cref="SentItemTriageMode"/> - a fixed action
        /// runs automatically; "Ask me each time" shows the prompt. Acts on the
        /// single sent item (conversation widening was dropped). Suppressed while
        /// we're acting on a previous choice (a move makes a copy transit Sent
        /// Items, which would otherwise re-trigger this handler on the transient
        /// copy).
        /// </summary>
        private void SentItems_ItemAdd(object item)
        {
            try
            {
                if (_suppressSentItemTriage) return;

                var mode = Settings.Load(_settingsStore).SentItemTriageMode;
                if (mode == SentItemTriageMode.Leave) return;

                var reference = _mailStore.ResolveMailItem(item);
                if (reference == null) return; // not a mail item (meeting response, report...)

                SentItemTriageAction action;
                if (mode == SentItemTriageMode.AskEveryTime)
                {
                    var triageVm = new SentItemTriageViewModel(reference.Subject);
                    new SentItemTriageWindow { DataContext = triageVm }.ShowDialog();
                    if (triageVm.SelectedAction == null || triageVm.SelectedAction == SentItemTriageAction.Leave)
                        return;
                    action = triageVm.SelectedAction.Value;
                }
                else
                {
                    action = mode == SentItemTriageMode.Delete
                        ? SentItemTriageAction.Delete
                        : SentItemTriageAction.MoveToInbox;
                }

                _suppressSentItemTriage = true;
                try { ApplySentItemTriage(action, reference); }
                finally { _suppressSentItemTriage = false; }
            }
            catch (Exception ex)
            {
                TryLog("Sent-item triage failed", ex);
            }
        }

        /// <summary>
        /// Apply a triage action to the single sent item (no conversation
        /// widening, no destination picker - those were dropped in the rework).
        /// </summary>
        private void ApplySentItemTriage(SentItemTriageAction action, MailItemRef item)
        {
            switch (action)
            {
                case SentItemTriageAction.Delete:
                    _mailStore.DeleteItem(item);
                    Log.Information("Sent-item triage deleted the sent mail.");
                    break;

                case SentItemTriageAction.MoveToInbox:
                    var inbox = _mailStore.GetInboxFolder();
                    if (inbox == null)
                    {
                        Log.Warning("Sent-item triage: could not resolve the Inbox folder.");
                        return;
                    }
                    var result = _classifier.Classify(
                        new ClassifyRequest(new[] { item }, new[] { inbox },
                                            keepCopy: false, removeAttachments: false,
                                            markTasksComplete: false,
                                            safetyCopy: _settingsStore.GetBool(
                                                SettingsKeys.ClassifySafetyCopy, false)));
                    Log.Information(
                        "Sent-item triage moved {Processed} mail(s) to the Inbox ({Errors} failed).",
                        result.ItemsProcessed, result.Errors);
                    break;
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

        private void TogglePane()
        {
            EnsureTaskPane();
            if (_taskPane == null) return;

            _taskPane.Visible = !_taskPane.Visible;
            if (_taskPane.Visible)
                TaskPaneServices.Host?.RefreshOnShow();
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

            // Read the persisted dock position; fall back to Right (= 2) if absent/invalid.
            int dockInt;
            if (!int.TryParse(_settingsStore?.Get(SettingsKeys.PaneDockPosition, null),
                              out dockInt))
                dockInt = (int)Office.MsoCTPDockPosition.msoCTPDockPositionRight;

            // CreateCTP instantiates the COM-registered host control by ProgId.
            _taskPane = _ctpFactory.CreateCTP(
                RblClassTaskPaneHost.ProgIdString, "RBLclass", Type.Missing);
            _taskPane.DockPosition = (Office.MsoCTPDockPosition)dockInt;
            _taskPane.Width = 360;
            _taskPane.DockPositionStateChange += OnTaskPaneDockPositionStateChange;
        }

        /// <summary>
        /// Persists the new dock position whenever the user moves the task pane.
        /// Fires on the Outlook UI (STA) thread.
        /// </summary>
        private void OnTaskPaneDockPositionStateChange(Office.CustomTaskPane customTaskPaneInst)
        {
            try
            {
                int pos = (int)customTaskPaneInst.DockPosition;
                _settingsStore?.Set(SettingsKeys.PaneDockPosition,
                    pos.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                TryLog("DockPositionStateChange handler failed", ex);
            }
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
                var loc = TaskPaneServices.Localization;
                var r = _lastIndexResult;
                int cached = _folderTree != null ? _folderTree.GetAll().Count : 0;

                string status;
                if (r == null)
                {
                    status = loc.GetString("MsgBox_IndexStatus_NotStarted");
                }
                else if (r.Source == IndexSource.NeedsWalk)
                {
                    status = loc.GetString("MsgBox_IndexStatus_WalkScheduled");
                }
                else
                {
                    status = loc.GetString("MsgBox_IndexStatus_Summary",
                        r.Source, r.StoreCount, r.FolderCount, cached);
                }

                string message = status + Environment.NewLine + Environment.NewLine +
                    loc.GetString("MsgBox_IndexStatus_Footer", _dbPath, _logDirectory);

                MessageBox.Show(message, loc.GetString("MsgBox_IndexStatus_Title"),
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError("Index status failed", ex);
            }
        }

        /// <summary>
        /// Ribbon "Refresh folders": re-walk the live stores on demand so folders
        /// created or renamed directly in Outlook (not via our own "New subfolder"
        /// action) surface in search. Reuses the first-run walk path
        /// (<see cref="IFolderTree.WalkAndPersist"/>); ribbon callbacks already run
        /// on the Outlook UI (STA) thread, which is where COM access must happen.
        /// </summary>
        public void OnRefreshFoldersClick(Office.IRibbonControl control)
        {
            try
            {
                var loc = TaskPaneServices.Localization;
                if (_folderTree == null)
                {
                    MessageBox.Show(loc.GetString("MsgBox_RefreshFolders_NotReady"),
                                    loc.GetString("MsgBox_RefreshFolders_Title"),
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Ignore a repeat click while a walk is already running.
                if (_folderTree.IndexStatus == IndexStatus.Indexing)
                    return;

                // Let the indicator paint 'Indexing' (yellow) before the
                // synchronous, STA-bound COM walk runs: WalkAndPersist sets the
                // status itself (Indexing -> Ready), and the dot - not a modal -
                // is the completion signal. The walk must stay on this (UI/STA)
                // thread because Outlook OM is single-threaded apartment.
                Dispatcher.CurrentDispatcher.BeginInvoke(
                    DispatcherPriority.Background,
                    new Action(RunFolderWalk));
            }
            catch (Exception ex)
            {
                ShowError("Refresh folders failed", ex);
            }
        }

        /// <summary>
        /// Runs a full folder walk on the Outlook UI (STA) thread and logs the
        /// result. Status transitions (Indexing -> Ready) are driven by
        /// <see cref="FolderIndexService.WalkAndPersist"/> and surfaced by the
        /// pane's colored indicator.
        /// </summary>
        private void RunFolderWalk()
        {
            try
            {
                var sw = Stopwatch.StartNew();
                _lastIndexResult = _folderTree.WalkAndPersist();
                sw.Stop();
                Log.Information(
                    "Folder refresh: {Stores} stores, {Folders} folders in {Ms} ms.",
                    _lastIndexResult.StoreCount, _lastIndexResult.FolderCount, sw.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                ShowError("Refresh folders failed", ex);
            }
        }

        /// <summary>Ribbon "Settings": modal dialog over every user-facing option (legacy §7). Live-applies on every change.</summary>
        public void OnSettingsClick(Office.IRibbonControl control)
        {
            try
            {
                new SettingsWindow
                {
                    DataContext = new SettingsViewModel(
                        _settingsStore,
                        () => _mailStore.GetSelectedItemHtmlBody())
                }.ShowDialog();
            }
            catch (Exception ex)
            {
                ShowError("Settings failed", ex);
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
                var loc = TaskPaneServices.Localization;
                string caption = loc != null ? loc.GetString("MsgBox_ErrorTitle", title) : "RBLclass - " + title;
                MessageBox.Show(ex.ToString(), caption,
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
