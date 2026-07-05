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
// 2.8.0.0 is a feedback wave: a pane "Save as" button and a standalone
// attachment-removal button, a "delete" action on the post-sent classify
// notification, and two classify-reliability fixes (auto-class vs. the
// attachment-disposition modal conflict, and the classify silent-failure/
// freeze investigation).
// Windows Installer ignores the 4th version field, so the bump to the third
// field is what the MSI's MajorUpgrade sees as newer and auto-removes the
// previous version.
[assembly: AssemblyVersion("2.8.0.0")]
[assembly: AssemblyFileVersion("2.8.0.0")]

// The add-in class opts into COM visibility explicitly with [ComVisible(true)];
// nothing else in the assembly should be exposed.
[assembly: ComVisible(false)]
