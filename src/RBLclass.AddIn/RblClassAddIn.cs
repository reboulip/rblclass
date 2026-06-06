using System;
using System.Diagnostics;
using System.IO;
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
    public class RblClassAddIn : IDTExtensibility2, Office.IRibbonExtensibility
    {
        // Kept in sync with release.config.json (Clsid / ProgId).
        public const string ClsidString = "43808654-547E-4222-9BE3-7FA4B781FA44";
        public const string ProgIdString = "RBLclass.AddIn";

        private OutlookOM.Application _outlookApp;

        private IFolderRepository _repository;
        private IMailStore _mailStore;
        private IFolderTree _folderTree;
        private IndexResult _lastIndexResult;

        // One-shot timer that runs the first-run folder walk on the Outlook UI
        // thread shortly AFTER startup, so the COM walk stays on the STA thread
        // (CLAUDE.md) without blocking Outlook's startup.
        private Timer _firstRunWalkTimer;

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

                _repository = new SqliteFolderRepository("Data Source=" + _dbPath);
                _mailStore = new OutlookMailStore(_outlookApp);
                _folderTree = new FolderIndexService(_mailStore, _repository);

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
