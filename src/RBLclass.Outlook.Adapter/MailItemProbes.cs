using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using RBLclass.Core;
using Serilog;
using OutlookOM = Microsoft.Office.Interop.Outlook;

namespace RBLclass.Outlook.Adapter
{
    /// <summary>
    /// Read-only probes of an opened mail item, shared by <see cref="OutlookMailStore"/>
    /// and <see cref="OutlookMailRebuilder"/> (moved verbatim from the store,
    /// v2.12.0.0 A5). STA only, like every COM access.
    /// </summary>
    internal static class MailItemProbes
    {
        internal const string PR_MESSAGE_CLASS_W = "http://schemas.microsoft.com/mapi/proptag/0x001A001F";
        internal const string PR_INTERNET_MESSAGE_ID_W = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
        internal const string DaslSubject = "urn:schemas:httpmail:subject";
        private const string PR_HASATTACH = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B";
        private const string PR_MESSAGE_SIZE = "http://schemas.microsoft.com/mapi/proptag/0x0E080003";

        /// <summary>
        /// True when the attachment is inline/embedded (a cid:-linked body image,
        /// signature logo, or embedded/OLE item) rather than a true detached file
        /// (v2.5.0.0 B2). Any one signal is sufficient:
        /// 1. Type is olEmbeddeditem or olOLE (embedded objects, never standalone files).
        /// 2. PR_ATTACHMENT_HIDDEN (0x7FFE000B) true - Outlook's hidden-attachment flag.
        /// 3. A content-id (PR_ATTACH_CONTENT_ID) AND the ATT_MHTML_REF bit in
        ///    PR_ATTACH_FLAGS - i.e. actually referenced from the HTML body. A
        ///    content-id ALONE is not enough: Outlook assigns one to ordinary
        ///    attachments too, so checking it alone wrongly hides real files.
        /// olByValue (normal files) and olByReference (linked files) are NOT inline.
        /// </summary>
        internal static bool IsInlineAttachment(OutlookOM.Attachment att)
        {
            var type = Safe(() => att.Type, OutlookOM.OlAttachmentType.olByValue);
            if (type == OutlookOM.OlAttachmentType.olEmbeddeditem
                || type == OutlookOM.OlAttachmentType.olOLE)
                return true;

            // The flags below need the PropertyAccessor, which not every store
            // exposes for every MAPI tag (PST vs Exchange differ), so failures
            // fall through to "not inline" (the attachment is shown, as pre-B2).
            try
            {
                using (var pa = new ComRef<OutlookOM.PropertyAccessor>(att.PropertyAccessor))
                {
                    const string PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";
                    if (Safe(() => (bool)pa.Value.GetProperty(PR_ATTACHMENT_HIDDEN), false))
                        return true;

                    const string PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";
                    const string PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003";
                    const int ATT_MHTML_REF = 0x00000004;
                    string cid = Safe(() => pa.Value.GetProperty(PR_ATTACH_CONTENT_ID) as string, null);
                    int flags = Safe(() => Convert.ToInt32(pa.Value.GetProperty(PR_ATTACH_FLAGS)), 0);
                    if (!string.IsNullOrEmpty(cid) && (flags & ATT_MHTML_REF) != 0)
                        return true;
                }
            }
            catch { /* property not available for this store/item - treat as not inline */ }

            return false;
        }

        /// <summary>
        /// The content id of a body image (v2.11.0.0 C2): a by-value attachment
        /// carrying PR_ATTACH_CONTENT_ID, read even when it is also hidden.
        /// Embedded items and OLE objects return null. Separate from
        /// <see cref="IsInlineAttachment"/>, which must keep reporting body images
        /// as inline so "remove all attachments" never takes a signature logo.
        /// </summary>
        internal static string ImageContentIdOf(OutlookOM.Attachment att)
        {
            if (Safe(() => att.Type, OutlookOM.OlAttachmentType.olOLE) != OutlookOM.OlAttachmentType.olByValue)
                return null;
            try
            {
                using (var pa = new ComRef<OutlookOM.PropertyAccessor>(att.PropertyAccessor))
                {
                    const string PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";
                    return InlineImageRules.NormalizeContentId(
                        Safe(() => pa.Value.GetProperty(PR_ATTACH_CONTENT_ID) as string, null));
                }
            }
            catch { return null; }
        }

        /// <summary>True when the item still shows its smime.p7m / smime.p7s part.</summary>
        internal static bool HasSmimePayloadAttachment(OutlookOM.MailItem mail)
        {
            OutlookOM.Attachments rawAtts;
            try { rawAtts = mail.Attachments; }
            catch { return false; }

            using (var attachments = new ComRef<OutlookOM.Attachments>(rawAtts))
            {
                int count = Safe(() => attachments.Value.Count, 0);
                for (int i = 1; i <= count; i++)
                {
                    OutlookOM.Attachment raw;
                    try { raw = attachments.Value[i]; } catch { continue; }
                    using (var att = new ComRef<OutlookOM.Attachment>(raw))
                    {
                        if (ProtectedMailPolicy.IsSmimePayloadName(Safe(() => att.Value.FileName, null)))
                            return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// The class the folder stores for this item when it is an S/MIME or secure
        /// class, else null (#62). Read from the parent folder's contents table,
        /// which is the stored row: an add-in that decrypts the opened item in
        /// memory does not change it. The table is restricted to protected classes
        /// and narrowed by Internet message id (or subject), so it holds a row or
        /// two; the row is matched with <c>CompareEntryIDs</c> since a table can
        /// return the short-term form of the EntryID. Best-effort: null on any fault.
        /// </summary>
        internal static string ReadStoredProtectedClass(OutlookOM.MailItem mail)
        {
            string entryId = Safe(() => mail.EntryID, null);
            if (string.IsNullOrEmpty(entryId)) return null; // never saved: nothing stored

            // No key to narrow on: skip, don't scan the folder.
            string filter = ProtectedRowFilter(ReadMessageId(mail), Safe(() => mail.Subject, null));
            if (filter == null) return null;

            object rawParent = null;
            try
            {
                rawParent = mail.Parent;
                var folder = rawParent as OutlookOM.Folder;
                if (folder == null) return null;

                using (var session = new ComRef<OutlookOM.NameSpace>(mail.Session))
                using (var table = new ComRef<OutlookOM.Table>(folder.GetTable(filter, OutlookOM.OlTableContents.olUserItems)))
                {
                    // Narrowed to one message id: a handful of rows at most (copies).
                    for (int guard = 0; guard < 25 && !table.Value.EndOfTable; guard++)
                    {
                        using (var row = new ComRef<OutlookOM.Row>(table.Value.GetNextRow()))
                        {
                            string rowId = Safe(() => row.Value["EntryID"] as string, null);
                            if (rowId == null) continue;
                            bool same = string.Equals(rowId, entryId, StringComparison.OrdinalIgnoreCase)
                                     || Safe(() => session.Value.CompareEntryIDs(rowId, entryId), false);
                            if (same) return Safe(() => row.Value["MessageClass"] as string, null);
                        }
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Stored message class lookup failed for {EntryId}.", entryId);
                return null;
            }
            finally
            {
                if (rawParent != null && Marshal.IsComObject(rawParent))
                    Marshal.ReleaseComObject(rawParent);
            }
        }

        /// <summary>The item's Internet message id, or null.</summary>
        internal static string ReadMessageId(OutlookOM.MailItem mail)
        {
            try
            {
                using (var pa = new ComRef<OutlookOM.PropertyAccessor>(mail.PropertyAccessor))
                {
                    string messageId = Safe(() => pa.Value.GetProperty(PR_INTERNET_MESSAGE_ID_W) as string, null);
                    return string.IsNullOrWhiteSpace(messageId) ? null : messageId;
                }
            }
            catch { return null; }
        }

        /// <summary>
        /// DASL filter for a folder table: protected classes only, narrowed on
        /// the message id, else on the subject. Null when there is neither.
        /// </summary>
        internal static string ProtectedRowFilter(string messageId, string subject)
        {
            string narrowing = RowNarrowing(messageId, subject);
            if (narrowing == null) return null;

            return "@SQL=(\"" + PR_MESSAGE_CLASS_W + "\" LIKE 'IPM.Note.SMIME%' OR \""
                 + PR_MESSAGE_CLASS_W + "\" LIKE 'IPM.Note.Secure%') AND " + narrowing;
        }

        /// <summary>DASL condition on the message id, else on the subject. Null when there is neither.</summary>
        private static string RowNarrowing(string messageId, string subject)
        {
            if (!string.IsNullOrWhiteSpace(messageId))
                return "\"" + PR_INTERNET_MESSAGE_ID_W + "\" = '" + DaslLiteral(messageId) + "'";
            if (!string.IsNullOrEmpty(subject))
                return "\"" + DaslSubject + "\" = '" + DaslLiteral(subject) + "'";
            return null;
        }

        /// <summary>What a folder's contents table shows for one stored message.</summary>
        internal sealed class StoredRow
        {
            public string MessageClass;
            public bool? HasAttachment;   // PR_HASATTACH, null when unreadable
            public long? Size;            // PR_MESSAGE_SIZE, null when unreadable
        }

        /// <summary>
        /// The stored class, attachment flag and size of message <paramref name="entryId"/>
        /// in <paramref name="folder"/>, read from the contents table narrowed on the
        /// message id (else the subject), any class (v2.14.0.1). No item is opened, so
        /// an add-in that decrypts opened items is not triggered. Null when the row is
        /// not found, there is no key to narrow on, or the table cannot be read.
        /// </summary>
        internal static StoredRow ReadStoredRow(OutlookOM.NameSpace session, OutlookOM.Folder folder,
                                                string entryId, string messageId, string subject)
        {
            string narrowing = RowNarrowing(messageId, subject);
            if (narrowing == null || folder == null || string.IsNullOrEmpty(entryId)) return null;
            try
            {
                using (var table = new ComRef<OutlookOM.Table>(folder.GetTable("@SQL=" + narrowing, OutlookOM.OlTableContents.olUserItems)))
                {
                    using (var columns = new ComRef<OutlookOM.Columns>(table.Value.Columns))
                    {
                        using (new ComRef<OutlookOM.Column>(columns.Value.Add(PR_HASATTACH))) { }
                        using (new ComRef<OutlookOM.Column>(columns.Value.Add(PR_MESSAGE_SIZE))) { }
                    }

                    // Narrowed to one message id or subject: a handful of rows.
                    for (int guard = 0; guard < 200 && !table.Value.EndOfTable; guard++)
                    {
                        using (var row = new ComRef<OutlookOM.Row>(table.Value.GetNextRow()))
                        {
                            string rowId = Safe(() => row.Value["EntryID"] as string, null);
                            if (rowId == null) continue;
                            bool same = string.Equals(rowId, entryId, StringComparison.OrdinalIgnoreCase)
                                     || (session != null && Safe(() => session.CompareEntryIDs(rowId, entryId), false));
                            if (!same) continue;

                            return new StoredRow
                            {
                                MessageClass = Safe(() => row.Value["MessageClass"] as string, null),
                                HasAttachment = Safe(() => (bool?)Convert.ToBoolean(row.Value[PR_HASATTACH]), null),
                                Size = Safe(() => (long?)Convert.ToInt64(row.Value[PR_MESSAGE_SIZE]), null),
                            };
                        }
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Stored row lookup failed for {EntryId}.", entryId);
                return null;
            }
        }

        /// <summary>
        /// The EntryIDs of the folder's stored protected rows matching the message
        /// id (or subject), read from the contents table: no item is opened, so an
        /// add-in that decrypts opened items is not triggered (v2.14.0.0 A2). Null
        /// when there is no key to narrow on or the table cannot be read.
        /// </summary>
        internal static List<string> FindProtectedEntryIds(OutlookOM.Folder folder, string messageId, string subject)
        {
            string filter = ProtectedRowFilter(messageId, subject);
            if (filter == null || folder == null) return null;
            try
            {
                var ids = new List<string>();
                using (var table = new ComRef<OutlookOM.Table>(folder.GetTable(filter, OutlookOM.OlTableContents.olUserItems)))
                {
                    // Narrowed to one message id or subject: a handful of rows.
                    for (int guard = 0; guard < 200 && !table.Value.EndOfTable; guard++)
                    {
                        using (var row = new ComRef<OutlookOM.Row>(table.Value.GetNextRow()))
                        {
                            string rowId = Safe(() => row.Value["EntryID"] as string, null);
                            if (!string.IsNullOrEmpty(rowId)) ids.Add(rowId);
                        }
                    }
                }
                return ids;
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Protected row lookup failed.");
                return null;
            }
        }

        /// <summary>Escape a value for a single-quoted DASL literal.</summary>
        internal static string DaslLiteral(string value) => value.Replace("'", "''");

        private static T Safe<T>(Func<T> read, T fallback)
        {
            try { return read(); }
            catch { return fallback; }
        }
    }
}