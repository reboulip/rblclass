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
    /// installer. It hosts WPF views through ElementHost and switches between the
    /// Open-folder and Classify views on demand.
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
        private FolderSearchView _openView;
        private ClassifyView _classifyView;

        public PaneMode CurrentMode { get; private set; } = PaneMode.OpenFolder;

        public RblClassTaskPaneHost()
        {
            try
            {
                // Publish ourselves so ribbon callbacks can switch our mode.
                TaskPaneServices.Host = this;

                _elementHost = new ElementHost { Dock = DockStyle.Fill };
                Controls.Add(_elementHost);

                ShowOpenFolder();
                Log.Information("TaskPaneHost initialized (searchWired={Wired}).",
                    TaskPaneServices.Search != null);
            }
            catch (Exception ex)
            {
                try { Log.Error(ex, "TaskPaneHost ctor failed."); } catch { }
            }
        }

        public void ShowOpenFolder()
        {
            if (_openView == null)
            {
                _openView = new FolderSearchView();
                if (TaskPaneServices.Search != null)
                    _openView.DataContext = new FolderSearchViewModel(
                        TaskPaneServices.Search, TaskPaneServices.Navigate, TaskPaneServices.Settings);
            }

            ApplyTheme(_openView);
            _elementHost.Child = _openView;
            CurrentMode = PaneMode.OpenFolder;
        }

        public void ShowClassify()
        {
            if (_classifyView == null)
            {
                _classifyView = new ClassifyView();
                if (TaskPaneServices.Search != null && TaskPaneServices.Classifier != null)
                    _classifyView.DataContext = new ClassifyViewModel(
                        TaskPaneServices.Search,
                        TaskPaneServices.Classifier,
                        TaskPaneServices.GetSelection,
                        TaskPaneServices.CreateSubfolder,
                        TaskPaneServices.Settings,
                        TaskPaneServices.ConfirmMarkTasksComplete);
            }

            // Re-read the live mail selection each time the pane is shown.
            (_classifyView.DataContext as ClassifyViewModel)?.RefreshSelection();

            ApplyTheme(_classifyView);
            _elementHost.Child = _classifyView;
            CurrentMode = PaneMode.Classify;
        }

        /// <summary>
        /// Theme the given view to the current Outlook look and match the
        /// WinForms host background so no white border shows around it. Called
        /// on each show, so a theme switch made while Outlook is running is
        /// picked up the next time the pane is opened.
        /// </summary>
        private void ApplyTheme(System.Windows.FrameworkElement view)
        {
            var mode = ThemeService.Apply(view);
            BackColor = ThemeService.WinFormsBackColor(mode);
            _elementHost.BackColor = BackColor;
        }

        /// <summary>Push a live selection count into the classify view model, if it exists.</summary>
        public void SetSelectionCount(int count)
        {
            (_classifyView?.DataContext as ClassifyViewModel)?.SetSelectionCount(count);
        }
    }
}
