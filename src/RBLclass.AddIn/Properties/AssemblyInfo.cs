using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("RBLclass.AddIn")]
[assembly: AssemblyDescription("RBLclass Outlook COM add-in - classify mail into PST folders with fast keyword search")]
[assembly: AssemblyProduct("RBLclass")]
[assembly: AssemblyCompany("RBLclass")]

// AssemblyVersion is pinned to 2.5.2.0 to match the strong assembly name in
// release.config.json's HKCU COM registration ("RBLclass.AddIn, Version=2.5.2.0,
// ..."). If these diverge, mscoree cannot resolve the add-in class at COM
// activation and Outlook silently fails to load it.
//
// 2.7.0.0 adds the v2.7.0.0 sprint: keyboard shortcuts (Ctrl+Z undo, open
// pane / focus search / open folder), post-send classify dialog clarity +
// vertical layout, configurable sent-items archive folder, multi-line
// search-result path display option, opt-in keep-latest-conversation on
// classify, update-notification investigation, and the end-user usage guide.
// Windows Installer ignores the 4th version field, so the bump to the third
// field is what the MSI's MajorUpgrade sees as newer and auto-removes the
// previous version.
[assembly: AssemblyVersion("2.7.0.0")]
[assembly: AssemblyFileVersion("2.7.0.0")]

// The add-in class opts into COM visibility explicitly with [ComVisible(true)];
// nothing else in the assembly should be exposed.
[assembly: ComVisible(false)]
