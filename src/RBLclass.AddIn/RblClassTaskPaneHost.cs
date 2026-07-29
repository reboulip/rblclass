using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using RBLclass.AddIn.Theming;
using RBLclass.AddIn.ViewModels;
using RBLclass.AddIn.Views;
using Serilog;

namespace RBLclass.AddIn
{
    /// <summary>
    /// The Windows Forms host control for the single RBLclass Custom Task Pane.
    /// Office instantiates it via COM (ICTPFactory.CreateCTP by ProgId), so it
    /// must be ComVisible with a stable [Guid]/[ProgId] matching
    /// release.config.json and be registered as an ActiveX control by the
    /// installer. It hosts one WPF view (<see cref="MainPaneView"/>) through
    /// ElementHost - the unified open-and-classify pane.
    /// </summary>
    [ComVisible(true)]
    [Guid(ClsidString)]
    [ProgId(ProgIdString)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class RblClassTaskPaneHost : UserControl
    {
        // Kept in sync with release.config.json TaskPaneControl.
        public const string ClsidString = "3357823B-8157-4B49-8B66-9BFF2F61C06B";
        public const string ProgIdString = "RBLclass.TaskPaneHost";

        private readonly ElementHost _elementHost;
        private MainPaneView _view;
        // Coalesces nav-rail inset recomputes queued from OnResize.
        private bool _railInsetPending;

        public RblClassTaskPaneHost()
        {
            try
            {
                // Publish ourselves so ribbon callbacks can show/refresh the pane.
                TaskPaneServices.Host = this;

                _elementHost = new ElementHost { Dock = DockStyle.Fill };
                Controls.Add(_elementHost);

                EnsureView();
                Log.Information("TaskPaneHost initialized (searchWired={Wired}).",
                    TaskPaneServices.Search != null);
            }
            catch (Exception ex)
            {
                try { Log.Error(ex, "TaskPaneHost ctor failed."); } catch { }
            }
        }

        private void EnsureView()
        {
            if (_view == null)
            {
                try
                {
                    _view = new MainPaneView();
                    if (TaskPaneServices.Search != null && TaskPaneServices.Classifier != null)
                    {
                        var vm = new MainPaneViewModel(
                            TaskPaneServices.Search,
                            TaskPaneServices.Classifier,
                            TaskPaneServices.GetSelection,
                            TaskPaneServices.CreateSubfolder,
                            TaskPaneServices.Navigate,
                            TaskPaneServices.Settings,
                            TaskPaneServices.ConfirmMarkTasksComplete,
                            TaskPaneServices.PromptForName,
                            TaskPaneServices.GetAllFolders,
                            TaskPaneServices.CheckHasMeetingItemsInSelection,
                            TaskPaneServices.ShowMeetingClassifyHint,
                            TaskPaneServices.PromptForRename,
                            TaskPaneServices.RenameFolder,
                            TaskPaneServices.CountFolderContents,
                            TaskPaneServices.ConfirmDeleteFolder,
                            TaskPaneServices.DeleteFolder);
                        vm.SubscribeToIndexStatus(TaskPaneServices.FolderIndex);
                        TaskPaneServices.PinMailForClassify = vm.PinMailForClassify;
                        _view.DataContext = vm;
                    }

                    _elementHost.Child = _view;
                }
                catch (Exception ex)
                {
                    // Reset so the next show retries construction instead of
                    // leaving a permanently blank pane (v2.7.1.0 B1).
                    try { Log.Error(ex, "EnsureView failed; will retry on next show."); } catch { }
                    _view = null;
                    return;
                }
            }

            ApplyThemeAndSelection();
        }

        /// <summary>
        /// Re-read the live mail selection and re-apply the theme. Called on each
        /// show so a theme switch or a changed selection is picked up.
        /// </summary>
        public void RefreshOnShow()
        {
            EnsureView();
            QueueNavRailInsetUpdate();
            _view?.FocusSearch();
        }

        /// <summary>
        /// Office gives a Top/Bottom-docked task pane the full Explorer client
        /// width, then paints the left navigation icon rail over its leftmost
        /// ~48px (carried-over v2.9.0.0 A2 finding; the classic CTP API cannot
        /// move the pane itself). Measure the covered strip and pad the WPF
        /// content right by that amount so nothing renders underneath the rail.
        /// Resize fires on every dock change, pane show, and Explorer resize,
        /// so the inset tracks the live geometry; it measures 0 (no padding)
        /// when docked Left/Right, floating, or when the rail is hidden.
        /// </summary>
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            QueueNavRailInsetUpdate();
        }

        private void QueueNavRailInsetUpdate()
        {
            if (_railInsetPending || !IsHandleCreated) return;
            _railInsetPending = true;
            try
            {
                BeginInvoke(new Action(() =>
                {
                    _railInsetPending = false;
                    try
                    {
                        var pad = new Padding(
                            NavRailInset.MeasureLeftOverlapPixels(Handle), 0, 0, 0);
                        if (Padding != pad)
                        {
                            Padding = pad;
                            Log.Information(
                                "Nav-rail inset applied: left padding {Pixels}px.",
                                pad.Left);
                        }
                    }
                    catch (Exception ex)
                    {
                        try { Log.Debug(ex, "Nav-rail inset measurement failed."); } catch { }
                    }
                }));
            }
            catch { _railInsetPending = false; }
        }

        /// <summary>Push a live selection count into the pane view model.</summary>
        public void SetSelectionCount(int count)
        {
            (_view?.DataContext as MainPaneViewModel)?.SetSelectionCount(count);
        }

        /// <summary>Easter egg: replace the Auto-Class button label with a pig.</summary>
        public void ActivatePigEasterEgg()
        {
            (_view?.DataContext as MainPaneViewModel)?.ActivatePigEasterEgg();
        }

        /// <summary>Show the opt-in "newer version available" banner (v2.7.0.0 E1).</summary>
        public void ShowUpdateBanner(string version)
        {
            (_view?.DataContext as MainPaneViewModel)?.ShowUpdateBanner(version);
        }

        /// <summary>
        /// Push a triage-outcome status message into the pane's rolling journal
        /// (v2.7.1.0 C1). Silently ignored when the pane is not yet initialized.
        /// </summary>
        public void PushExternalStatus(string message)
        {
            (_view?.DataContext as MainPaneViewModel)?.PushExternalStatus(message);
        }

        /// <summary>
        /// Re-seed the pane option defaults from settings after the Settings
        /// dialog closes (v2.7.1.0 D3).
        /// </summary>
        public void ReloadOptionDefaults()
        {
            (_view?.DataContext as MainPaneViewModel)?.ReloadOptionDefaults();
        }

        private void ApplyThemeAndSelection()
        {
            (_view?.DataContext as MainPaneViewModel)?.RefreshSelection();

            var mode = ThemeService.Apply(_view);
            BackColor = ThemeService.WinFormsBackColor(mode);
            _elementHost.BackColor = BackColor;
        }
    }
}
