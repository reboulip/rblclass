using System;
using System.Runtime.InteropServices;

namespace RBLclass.HelloPstPoc
{
    // Manual declarations of the Office interop interfaces we need.
    // We declare them by hand instead of taking COM/NuGet references
    // on Extensibility 1.3 and the Microsoft Office object library
    // (no maintained NuGet package exists for the latter), so the
    // project builds against nothing but Microsoft.Office.Interop.Outlook.

    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        // DispIds must match the canonical Extensibility 1.3 type
        // library exactly — Outlook calls IDispatch::Invoke with these
        // hard-coded values. Without explicit DispIds, .NET auto-
        // assigns and Outlook's call dispatches to the wrong method.
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object application,
            ext_ConnectMode connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInst,
            ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            ext_DisconnectMode removeMode,
            ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(ref Array custom);

        [DispId(4)]
        void OnStartupComplete(ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(ref Array custom);
    }

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5,
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
    }

    // Office.IRibbonExtensibility — Outlook queries the add-in for this
    // interface by GUID at startup to fetch the ribbon XML.
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }
}
