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
// 2.6.0.0 adds the v2.6.0.0 sprint: ribbon UI overhaul (Settings/About/easter
// egg), Classify button prominence, search auto-expand + row separator,
// auto-classify history retention setting, same-save-dir attachment checkbox,
// external-sender banner diagnostics + config, and MeetingItem classify support.
// Windows Installer ignores the 4th version field, so the bump to the third
// field is what the MSI's MajorUpgrade sees as newer and auto-removes the
// previous version.
[assembly: AssemblyVersion("2.6.0.0")]
[assembly: AssemblyFileVersion("2.6.0.0")]

// The add-in class opts into COM visibility explicitly with [ComVisible(true)];
// nothing else in the assembly should be exposed.
[assembly: ComVisible(false)]
