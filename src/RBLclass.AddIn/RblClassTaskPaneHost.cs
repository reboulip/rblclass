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
                _view = new MainPaneView();
                if (TaskPaneServices.Search != null && TaskPaneServices.Classifier != null)
                    _view.DataContext = new MainPaneViewModel(
                        TaskPaneServices.Search,
                        TaskPaneServices.Classifier,
                        TaskPaneServices.GetSelection,
                        TaskPaneServices.CreateSubfolder,
                        TaskPaneServices.Navigate,
                        TaskPaneServices.Settings,
                        TaskPaneServices.ConfirmMarkTasksComplete,
                        TaskPaneServices.PromptForName);

                _elementHost.Child = _view;
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
        }

        /// <summary>Push a live selection count into the pane view model.</summary>
        public void SetSelectionCount(int count)
        {
            (_view?.DataContext as MainPaneViewModel)?.SetSelectionCount(count);
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
