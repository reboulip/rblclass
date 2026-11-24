using System;
using System.Globalization;
using System.Runtime.InteropServices;
using Serilog;
using OutlookOM = Microsoft.Office.Interop.Outlook;

namespace RBLclass.Outlook.Adapter
{
    /// <summary>
    /// Store-level move or copy of one message through Extended MAPI
    /// <c>IMAPIFolder::CopyMessages</c> (v2.14.0.0 A1, GH #62): the call Outlook
    /// makes on a drag-and-drop. It copies the stored message, its class and its
    /// <c>smime.p7m</c> part byte for byte, without opening it through the object
    /// model, so an add-in that decrypts opened items in memory (Stormshield) is
    /// never involved and no decrypted state is written.
    /// </summary>
    /// <remarks>
    /// <para>
    /// No <c>[ComImport]</c> declaration: CLAUDE.md forbids hand-rolled interface
    /// declarations because the CLR builds an IL stub for every method and one
    /// wrong signature takes Outlook down. There is no PIA for MAPI, so the single
    /// method used is called through its vtable slot with one explicit delegate.
    /// </para>
    /// <para>
    /// <c>Folder.MAPIOBJECT</c> is only usable in process (inside Outlook) and on
    /// the Outlook UI thread (STA), like every other adapter call.
    /// </para>
    /// </remarks>
    internal static class MapiMessageMover
    {
        private static readonly Guid IID_IMAPIFolder = new Guid("0002030C-0000-0000-C000-000000000046");

        /// <summary>
        /// IUnknown 0-2; IMAPIProp 3-13 (GetLastError, SaveChanges, GetProps,
        /// GetPropList, OpenProperty, SetProps, DeleteProps, CopyTo, CopyProps,
        /// GetNamesFromIDs, GetIDsFromNames); IMAPIContainer 14-18
        /// (GetContentsTable, GetHierarchyTable, OpenEntry, SetSearchCriteria,
        /// GetSearchCriteria); IMAPIFolder 19 CreateMessage, 20 CopyMessages.
        /// </summary>
        private const int CopyMessagesSlot = 20;

        /// <summary>IMAPIFolder 21 DeleteMessages (right after CopyMessages).</summary>
        private const int DeleteMessagesSlot = 21;

        private const uint MESSAGE_MOVE = 0x00000001;
        private const int MAPI_W_PARTIAL_COMPLETION = 0x00040680;

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int CopyMessagesFn(IntPtr self, IntPtr lpMsgList, IntPtr lpInterface,
                                            IntPtr lpDestFolder, IntPtr ulUIParam, IntPtr lpProgress,
                                            uint ulFlags);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int DeleteMessagesFn(IntPtr self, IntPtr lpMsgList, IntPtr ulUIParam,
                                              IntPtr lpProgress, uint ulFlags);

        /// <summary>MAPI <c>SBinary</c>: the IntPtr field gives the 64-bit padding.</summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct SBinary
        {
            public uint cb;
            public IntPtr lpb;
        }

        /// <summary>MAPI <c>ENTRYLIST</c> (an <c>SBinaryArray</c>).</summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct EntryList
        {
            public uint cValues;
            public IntPtr lpbin;
        }

        /// <summary>
        /// Move (<paramref name="move"/> true) or copy the message
        /// <paramref name="entryIdHex"/> from <paramref name="source"/> into
        /// <paramref name="destination"/>. Never throws: false on any failure,
        /// logged with the HRESULT.
        /// </summary>
        internal static bool TryTransfer(OutlookOM.Folder source, OutlookOM.Folder destination,
                                         string entryIdHex, bool move)
        {
            if (source == null || destination == null) return false;
            string op = move ? "move" : "copy";
            IntPtr dst = IntPtr.Zero;
            try
            {
                dst = GetFolderInterface(destination);
                if (dst == IntPtr.Zero)
                {
                    Log.Warning("Store-level filing: the destination exposes no IMAPIFolder ({EntryId}).", entryIdHex);
                    return false;
                }
                return Invoke(source, entryIdHex, "CopyMessages", op, (src, list) =>
                {
                    var copyMessages = Method<CopyMessagesFn>(src, CopyMessagesSlot);
                    // lpInterface null = IMAPIFolder for the destination; no UI, no progress.
                    int hr = copyMessages(src, list, IntPtr.Zero, dst, IntPtr.Zero, IntPtr.Zero,
                                          move ? MESSAGE_MOVE : 0u);
                    GC.KeepAlive(copyMessages);
                    return hr;
                });
            }
            finally
            {
                if (dst != IntPtr.Zero) Marshal.Release(dst);
            }
        }

        /// <summary>
        /// Hard-delete the message <paramref name="entryIdHex"/> from
        /// <paramref name="folder"/> (no Deleted Items step, nothing opened). Never
        /// throws: false on any failure, logged with the HRESULT.
        /// </summary>
        internal static bool TryDelete(OutlookOM.Folder folder, string entryIdHex)
        {
            if (folder == null) return false;
            return Invoke(folder, entryIdHex, "DeleteMessages", "delete", (src, list) =>
            {
                var deleteMessages = Method<DeleteMessagesFn>(src, DeleteMessagesSlot);
                int hr = deleteMessages(src, list, IntPtr.Zero, IntPtr.Zero, 0u);
                GC.KeepAlive(deleteMessages);
                return hr;
            });
        }

        /// <summary>The vtable method at <paramref name="slot"/> of a MAPI interface pointer.</summary>
        private static T Method<T>(IntPtr self, int slot) where T : class
        {
            IntPtr vtable = Marshal.ReadIntPtr(self);
            IntPtr fn = Marshal.ReadIntPtr(vtable, slot * IntPtr.Size);
            return (T)(object)Marshal.GetDelegateForFunctionPointer(fn, typeof(T));
        }

        /// <summary>
        /// Build a one-entry ENTRYLIST, get the folder's IMAPIFolder, run
        /// <paramref name="call"/> (folder pointer, list pointer) -> HRESULT, and
        /// free everything. False on a failure HRESULT or a fault.
        /// </summary>
        private static bool Invoke(OutlookOM.Folder folder, string entryIdHex, string method, string op,
                                   Func<IntPtr, IntPtr, int> call)
        {
            byte[] entryId = EntryIdBytes(entryIdHex);
            if (entryId == null)
            {
                Log.Warning("Store-level {Op}: the entry ID is not valid hex ({EntryId}).", op, entryIdHex);
                return false;
            }

            IntPtr src = IntPtr.Zero, bytes = IntPtr.Zero, bin = IntPtr.Zero, list = IntPtr.Zero;
            try
            {
                src = GetFolderInterface(folder);
                if (src == IntPtr.Zero)
                {
                    Log.Warning("Store-level {Op}: the folder exposes no IMAPIFolder ({EntryId}).", op, entryIdHex);
                    return false;
                }

                bytes = Marshal.AllocHGlobal(entryId.Length);
                Marshal.Copy(entryId, 0, bytes, entryId.Length);

                bin = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(SBinary)));
                Marshal.StructureToPtr(new SBinary { cb = (uint)entryId.Length, lpb = bytes }, bin, false);

                list = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(EntryList)));
                Marshal.StructureToPtr(new EntryList { cValues = 1, lpbin = bin }, list, false);

                int hr = call(src, list);
                if (hr < 0 || hr == MAPI_W_PARTIAL_COMPLETION)
                {
                    Log.Warning("Store-level {Op}: {Method} failed (hr=0x{Hr:X8}, {EntryId}).", op, method, hr, entryIdHex);
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Store-level {Op}: {Method} threw ({EntryId}).", op, method, entryIdHex);
                return false;
            }
            finally
            {
                if (list != IntPtr.Zero) Marshal.FreeHGlobal(list);
                if (bin != IntPtr.Zero) Marshal.FreeHGlobal(bin);
                if (bytes != IntPtr.Zero) Marshal.FreeHGlobal(bytes);
                if (src != IntPtr.Zero) Marshal.Release(src);
            }
        }

        /// <summary>The folder's AddRef'd IMAPIFolder pointer, or zero.</summary>
        private static IntPtr GetFolderInterface(OutlookOM.Folder folder)
        {
            object mapi = null;
            IntPtr unk = IntPtr.Zero;
            try
            {
                mapi = folder.MAPIOBJECT;
                if (mapi == null) return IntPtr.Zero;
                unk = Marshal.GetIUnknownForObject(mapi);
                Guid iid = IID_IMAPIFolder;
                int hr = Marshal.QueryInterface(unk, ref iid, out IntPtr ptr);
                return hr >= 0 ? ptr : IntPtr.Zero;
            }
            finally
            {
                if (unk != IntPtr.Zero) Marshal.Release(unk);
                if (mapi != null && Marshal.IsComObject(mapi)) Marshal.ReleaseComObject(mapi);
            }
        }

        /// <summary>Hex EntryID (as Outlook prints it) to bytes; null when malformed.</summary>
        internal static byte[] EntryIdBytes(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex.Length % 2 != 0) return null;
            var bytes = new byte[hex.Length / 2];
            for (int i = 0; i < bytes.Length; i++)
            {
                if (!byte.TryParse(hex.Substring(i * 2, 2), NumberStyles.HexNumber,
                                   CultureInfo.InvariantCulture, out bytes[i]))
                    return null;
            }
            return bytes;
        }
    }
}
