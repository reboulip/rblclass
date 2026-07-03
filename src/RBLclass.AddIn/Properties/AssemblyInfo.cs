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
// 2.7.1.0 is the code-review remediation sprint: S/MIME-safe mark-as-read,
// attachment-disposition and keep-copy filing integrity, exception guards on
// all UI entry points, classify/auto-classify error accounting, sent-item
// triage hardening, gated startup pane auto-show, immediate settings
// propagation, settings-store WAL + batch reads, and search hot-path caching.
// Windows Installer ignores the 4th version field, so the bump to the third
// field is what the MSI's MajorUpgrade sees as newer and auto-removes the
// previous version.
[assembly: AssemblyVersion("2.7.1.0")]
[assembly: AssemblyFileVersion("2.7.1.0")]

// The add-in class opts into COM visibility explicitly with [ComVisible(true)];
// nothing else in the assembly should be exposed.
[assembly: ComVisible(false)]
