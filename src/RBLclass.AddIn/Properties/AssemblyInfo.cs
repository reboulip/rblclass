using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("RBLclass.AddIn")]
[assembly: AssemblyDescription("RBLclass Outlook COM add-in - classify mail into PST folders with fast keyword search")]
[assembly: AssemblyProduct("RBLclass")]
[assembly: AssemblyCompany("RBLclass")]

// AssemblyVersion is pinned to match the strong assembly name in
// release.config.json's HKCU COM registration ("RBLclass.AddIn, Version=2.9.1.0,
// ..."). If these diverge, mscoree cannot resolve the add-in class at COM
// activation and Outlook silently fails to load it.
//
// 2.9.1.0 is a hotfix: the Top/Bottom live-redock freeze (CustomTaskPane
// docking manager wedged by a post-show Width/Height set) plus the nav-rail
// content overlap it was found alongside. Windows Installer ignores the 4th
// version field, so the bump to the third field is what the MSI's
// MajorUpgrade sees as newer and auto-removes the previous version.
[assembly: AssemblyVersion("2.9.1.0")]
[assembly: AssemblyFileVersion("2.9.1.0")]

// The add-in class opts into COM visibility explicitly with [ComVisible(true)];
// nothing else in the assembly should be exposed.
[assembly: ComVisible(false)]
