using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using RBLclass.AddIn.ViewModels;
using RBLclass.AddIn.Views;
using Serilog;

namespace RBLclass.AddIn
{
    /// <summary>
    /// The Windows Forms host control for the RBLclass Custom Task Pane. Office
    /// instantiates it via COM (ICTPFactory.CreateCTP by ProgId), so it must be
    /// ComVisible with a stable [Guid]/[ProgId] matching release.config.json and
    /// be registered as an ActiveX control by the installer. It hosts the WPF
    /// folder-search view through an ElementHost (WPF-in-WinForms bridge).
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

        public RblClassTaskPaneHost()
        {
            try
            {
                var elementHost = new ElementHost { Dock = DockStyle.Fill };
                var view = new FolderSearchView();

                if (TaskPaneServices.Search != null)
                    view.DataContext = new FolderSearchViewModel(
                        TaskPaneServices.Search,
                        TaskPaneServices.Navigate,
                        TaskPaneServices.Settings);

                elementHost.Child = view;
                Controls.Add(elementHost);

                Log.Information("TaskPaneHost initialized (searchWired={Wired}).",
                    TaskPaneServices.Search != null);
            }
            catch (Exception ex)
            {
                try { Log.Error(ex, "TaskPaneHost ctor failed."); } catch { }
            }
        }
    }
}
