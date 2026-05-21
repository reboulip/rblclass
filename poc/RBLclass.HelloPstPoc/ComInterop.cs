using System;
using System.Runtime.InteropServices;

namespace RBLclass.HelloPstPoc
{
    // Manual declaration of the Office Extensibility IDTExtensibility2
    // interface and its enums. We do this instead of taking a COM
    // reference on the Extensibility 1.3 type library so the project
    // builds without an external interop dependency.

    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object application,
            ext_ConnectMode connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInst,
            ref Array custom);

        void OnDisconnection(
            ext_DisconnectMode removeMode,
            ref Array custom);

        void OnAddInsUpdate(ref Array custom);

        void OnStartupComplete(ref Array custom);

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
}
