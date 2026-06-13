using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("RBLclass.AddIn")]
[assembly: AssemblyDescription("RBLclass Outlook COM add-in - classify mail into PST folders with fast keyword search")]
[assembly: AssemblyProduct("RBLclass")]
[assembly: AssemblyCompany("RBLclass")]

// AssemblyVersion is pinned to 2.3.0.0 to match the strong assembly name in
// release.config.json's HKCU COM registration ("RBLclass.AddIn, Version=2.3.0.0,
// ..."). If these diverge, mscoree cannot resolve the add-in class at COM
// activation and Outlook silently fails to load it.
[assembly: AssemblyVersion("2.3.0.0")]
[assembly: AssemblyFileVersion("2.3.0.0")]

// The add-in class opts into COM visibility explicitly with [ComVisible(true)];
// nothing else in the assembly should be exposed.
[assembly: ComVisible(false)]
