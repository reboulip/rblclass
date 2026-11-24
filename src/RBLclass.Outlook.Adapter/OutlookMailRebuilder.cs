using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using RBLclass.Core;
using Serilog;
using OutlookOM = Microsoft.Office.Interop.Outlook;

namespace RBLclass.Outlook.Adapter
{
    /// <summary>
    /// <see cref="IMailRebuilder"/> over the Outlook Object Model and the in-box
    /// .NET Framework CMS implementation (<c>System.Security.dll</c>), v2.12.0.0
    /// A5 (experimental re-encrypted filing, GH #62).
    /// </summary>
    /// <remarks>
    /// COM rules from CLAUDE.md apply: every COM object in a <see cref="ComRef{T}"/>,
    /// indexed loops, no chained property access, STA only. Logs never carry a
    /// subject, body, file name or address: tags, OIDs, counts, thumbprints and
    /// EntryIDs only.
    /// </remarks>
    public sealed class OutlookMailRebuilder : IMailRebuilder
    {
        /// <summary>Total attachment bytes above which an item is not rebuilt (32-bit memory budget).</summary>
        public const long MaxContentBytes = 25L * 1024 * 1024;

        private const string OidData = "1.2.840.113549.1.7.1";
        private const string OidAes256Cbc = "2.16.840.1.101.3.4.1.42";
        private const string OidRsa = "1.2.840.113549.1.1.1";
        private const string OidSan = "2.5.29.17";

        private const string Tag = "http://schemas.microsoft.com/mapi/proptag/";
        private const string PR_ATTACH_DATA_BIN = Tag + "0x37010102";
        private const string PR_ATTACH_MIME_TAG_W = Tag + "0x370E001F";
        private const string PR_SMTP_ADDRESS_W = Tag + "0x39FE001F";
        private const string PR_DISPLAY_NAME_W = Tag + "0x3001001F";
        private const string PR_MESSAGE_DELIVERY_TIME = Tag + "0x0E060040";
        private const string PR_ICON_INDEX = Tag + "0x10800003";

        /// <summary>
        /// Envelope properties copied from the source to the rebuilt item. The
        /// identity tags are deleted on the copy when the source lacks them: a
        /// PostItem stamps the current user, and Reply would then go to yourself.
        /// </summary>
        private static readonly string[] EnvelopeTags =
        {
            Tag + "0x0C1A001F", // PR_SENDER_NAME_W
            Tag + "0x0C1F001F", // PR_SENDER_EMAIL_ADDRESS_W
            Tag + "0x0C1E001F", // PR_SENDER_ADDRTYPE_W
            Tag + "0x0C190102", // PR_SENDER_ENTRYID
            Tag + "0x0C1D0102", // PR_SENDER_SEARCH_KEY
            Tag + "0x5D01001F", // PR_SENDER_SMTP_ADDRESS_W
            Tag + "0x0042001F", // PR_SENT_REPRESENTING_NAME_W
            Tag + "0x0065001F", // PR_SENT_REPRESENTING_EMAIL_ADDRESS_W
            Tag + "0x0064001F", // PR_SENT_REPRESENTING_ADDRTYPE_W
            Tag + "0x00410102", // PR_SENT_REPRESENTING_ENTRYID
            Tag + "0x003B0102", // PR_SENT_REPRESENTING_SEARCH_KEY
            Tag + "0x5D02001F", // PR_SENT_REPRESENTING_SMTP_ADDRESS_W
            Tag + "0x00390040", // PR_CLIENT_SUBMIT_TIME
            PR_MESSAGE_DELIVERY_TIME,
            Tag + "0x1035001F", // PR_INTERNET_MESSAGE_ID_W
            Tag + "0x0070001F", // PR_CONVERSATION_TOPIC_W
            Tag + "0x00710102", // PR_CONVERSATION_INDEX
            Tag + "0x10900003", // PR_FLAG_STATUS
        };

        private readonly OutlookOM.Application _app;
        private readonly string _tempDirectory;

        /// <param name="tempDirectory">
        /// Where an attachment too large for <c>PropertyAccessor</c> is staged
        /// (zeroed then deleted) and where the ciphertext payload is written to be
        /// attached. Created on demand.
        /// </param>
        public OutlookMailRebuilder(OutlookOM.Application app, string tempDirectory)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _tempDirectory = tempDirectory ?? throw new ArgumentNullException(nameof(tempDirectory));
        }

        // ---- read ----------------------------------------------------------------

        public MailContent ReadDecryptedContent(MailItemRef item)
        {
            if (item == null) return null;
            try
            {
                using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
                using (var comItem = new ComRef<object>(session.Value.GetItemFromID(item.EntryId, item.StoreId)))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return Refuse(item, "not a mail item");

                    if (MailItemProbes.HasSmimePayloadAttachment(mail))
                        return Refuse(item, "S/MIME payload still visible (not decrypted)");

                    string liveClass = Safe(() => mail.MessageClass, string.Empty);
                    if (IsSignedOnlyClass(liveClass) || IsSignedOnlyClass(MailItemProbes.ReadStoredProtectedClass(mail)))
                        return Refuse(item, "signed only");

                    var parts = ReadParts(mail, item);
                    if (parts == null) return null;

                    bool isHtml = Safe(() => mail.BodyFormat, OutlookOM.OlBodyFormat.olFormatUnspecified)
                                  == OutlookOM.OlBodyFormat.olFormatHTML;
                    string html = isHtml ? Safe(() => mail.HTMLBody, null) : null;
                    string plain = isHtml ? null : Safe(() => mail.Body, null);
                    return new MailContent(html, plain, parts);
                }
            }
            catch (OutOfMemoryException)
            {
                return Refuse(item, "out of memory");
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Re-encryption: reading the decrypted content failed ({EntryId}).", item.EntryId);
                return null;
            }
        }

        private List<MailPart> ReadParts(OutlookOM.MailItem mail, MailItemRef item)
        {
            using (var attachments = new ComRef<OutlookOM.Attachments>(mail.Attachments))
            {
                int count = attachments.Value.Count;

                // Size pre-check before reading any byte.
                long declared = 0;
                for (int i = 1; i <= count; i++)
                    using (var att = new ComRef<OutlookOM.Attachment>(attachments.Value[i]))
                        declared += Safe(() => (long)att.Value.Size, 0L);
                if (declared > MaxContentBytes)
                {
                    Refuse(item, "over the size cap (" + declared + " bytes)");
                    return null;
                }

                var parts = new List<MailPart>(count);
                long total = 0;
                for (int i = 1; i <= count; i++)
                {
                    using (var att = new ComRef<OutlookOM.Attachment>(attachments.Value[i]))
                    {
                        var type = att.Value.Type;
                        if (type != OutlookOM.OlAttachmentType.olByValue)
                        {
                            // Embedded items, OLE objects and links cannot be rebuilt faithfully.
                            Refuse(item, "attachment type " + type);
                            return null;
                        }

                        byte[] data = ReadAttachmentBytes(att.Value);
                        if (data == null)
                        {
                            Refuse(item, "an attachment could not be read");
                            return null;
                        }
                        total += data.Length;
                        if (total > MaxContentBytes)
                        {
                            Refuse(item, "over the size cap (" + total + " bytes read)");
                            return null;
                        }

                        string fileName = Safe(() => att.Value.FileName, string.Empty);
                        bool inline = MailItemProbes.IsInlineAttachment(att.Value);
                        // Read for every part: an unflagged body image keeps its cid.
                        string contentId = MailItemProbes.ImageContentIdOf(att.Value);
                        int index = Safe(() => att.Value.Index, i);
                        parts.Add(new MailPart(index, fileName, ReadMimeTag(att.Value, fileName), data, inline, contentId));
                    }
                }
                return parts;
            }
        }

        private static string ReadMimeTag(OutlookOM.Attachment att, string fileName)
        {
            try
            {
                using (var pa = new ComRef<OutlookOM.PropertyAccessor>(att.PropertyAccessor))
                {
                    string tag = Safe(() => pa.Value.GetProperty(PR_ATTACH_MIME_TAG_W) as string, null);
                    if (!string.IsNullOrWhiteSpace(tag)) return tag;
                }
            }
            catch { /* no tag: guess below */ }

            // An inline image declared octet-stream may not render: map the common ones.
            switch ((Path.GetExtension(fileName ?? string.Empty) ?? string.Empty).ToLowerInvariant())
            {
                case ".png": return "image/png";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".gif": return "image/gif";
                case ".bmp": return "image/bmp";
                default: return null;
            }
        }

        /// <summary>
        /// The attachment's bytes: through <c>PropertyAccessor</c> when it allows
        /// it (no disk), else staged with <c>SaveAsFile</c> in the temp directory,
        /// read, zeroed and deleted. Null on failure.
        /// </summary>
        private byte[] ReadAttachmentBytes(OutlookOM.Attachment att)
        {
            try
            {
                using (var pa = new ComRef<OutlookOM.PropertyAccessor>(att.PropertyAccessor))
                {
                    var bytes = pa.Value.GetProperty(PR_ATTACH_DATA_BIN) as byte[];
                    if (bytes != null) return bytes;
                }
            }
            catch { /* too large for GetProps on this store: stage on disk */ }

            string path = null;
            try
            {
                Directory.CreateDirectory(_tempDirectory);
                path = Path.Combine(_tempDirectory, Guid.NewGuid().ToString("N"));
                att.SaveAsFile(path);
                Log.Debug("Re-encryption: an attachment was staged on disk (PropertyAccessor refused it).");
                return File.ReadAllBytes(path);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Re-encryption: an attachment could not be read.");
                return null;
            }
            finally
            {
                if (path != null) ScrubAndDelete(path);
            }
        }

        private static MailContent Refuse(MailItemRef item, string reason)
        {
            Log.Information("Re-encryption: item not rebuilt ({Reason}, {EntryId}).", reason, item.EntryId);
            return null;
        }

        private static bool IsSignedOnlyClass(string messageClass) =>
            !string.IsNullOrEmpty(messageClass)
            && (messageClass.StartsWith("IPM.Note.SMIME.MultipartSigned", StringComparison.OrdinalIgnoreCase)
                || messageClass.StartsWith("IPM.Note.Secure.Sign", StringComparison.OrdinalIgnoreCase));

        // ---- encrypt ---------------------------------------------------------------

        public byte[] Encrypt(byte[] mimeEntity, string certificateThumbprint)
        {
            if (mimeEntity == null) return null;
            X509Certificate2 certificate = null;
            try
            {
                certificate = FindCertificate(certificateThumbprint);
                if (certificate == null)
                {
                    Log.Warning("Re-encryption: certificate {Thumbprint} is not a valid candidate.", certificateThumbprint);
                    return null;
                }

                // v2.14.0.0 B1: the envelope Stormshield uses (RSA PKCS#1 v1.5 key
                // transport). EnvelopedCms on .NET Framework 4.8 emits RSA-OAEP and
                // has no padding option; a target crashed on those copies.
                RSAParameters publicKey;
                using (var rsa = certificate.GetRSAPublicKey())
                    publicKey = rsa.ExportParameters(false);
                byte[] serial = certificate.GetSerialNumber();   // little-endian
                Array.Reverse(serial);
                byte[] payload = CmsEnvelopeBuilder.Build(mimeEntity, certificate.IssuerName.RawData, serial, publicKey);

                // Self-check: decode the envelope (no decryption) and record what was produced.
                var check = new EnvelopedCms();
                check.Decode(payload);
                string cipher = check.ContentEncryptionAlgorithm.Oid.Value;
                string keyTransport = check.RecipientInfos.Count > 0
                    ? check.RecipientInfos[0].KeyEncryptionAlgorithm.Oid.Value
                    : "none";
                Log.Debug("Re-encryption: payload {Bytes} bytes, cipher {Cipher}, key transport {KeyTransport}.",
                          payload.Length, cipher, keyTransport);
                if (cipher != OidAes256Cbc)
                    Log.Warning("Re-encryption: unexpected content cipher {Cipher}.", cipher);
                if (keyTransport != OidRsa)
                    Log.Warning("Re-encryption: unexpected key transport {KeyTransport}.", keyTransport);
                return payload;
            }
            catch (OutOfMemoryException)
            {
                Log.Warning("Re-encryption: out of memory while encrypting.");
                return null;
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Re-encryption: encryption failed.");
                return null;
            }
            finally
            {
                certificate?.Dispose();
            }
        }

        private static X509Certificate2 FindCertificate(string thumbprint)
        {
            string wanted = EncryptionCertificatePolicy.NormalizeThumbprint(thumbprint);
            if (wanted.Length == 0) return null;

            var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            try
            {
                store.Open(OpenFlags.ReadOnly);
                foreach (var cert in store.Certificates)
                {
                    if (EncryptionCertificatePolicy.NormalizeThumbprint(cert.Thumbprint) == wanted
                        && EncryptionCertificatePolicy.IsCandidate(ToCandidate(cert), DateTime.Now))
                        return cert;
                    cert.Dispose();
                }
                return null;
            }
            finally { store.Close(); }
        }

        // ---- certificates ----------------------------------------------------------

        public IReadOnlyList<CertificateInfo> ListEncryptionCertificates()
        {
            try
            {
                return EncryptionCertificatePolicy.Candidates(ReadStoreCandidates(), DateTime.Now)
                                                  .Select(c => c.ToInfo()).ToList();
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Re-encryption: listing certificates failed.");
                return new CertificateInfo[0];
            }
        }

        public CertificateInfo ResolveCertificate(string preferredThumbprint)
        {
            try
            {
                var all = ReadStoreCandidates();
                var chosen = string.IsNullOrWhiteSpace(preferredThumbprint)
                    ? EncryptionCertificatePolicy.SelectAutomatic(all, ReadAccountSmtp(), DateTime.Now)
                    : EncryptionCertificatePolicy.ResolveExplicit(all, preferredThumbprint, DateTime.Now);
                return chosen?.ToInfo();
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Re-encryption: resolving the certificate failed.");
                return null;
            }
        }

        private static List<CertificateCandidate> ReadStoreCandidates()
        {
            var result = new List<CertificateCandidate>();
            var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            try
            {
                store.Open(OpenFlags.ReadOnly);
                foreach (var cert in store.Certificates)
                {
                    using (cert) result.Add(ToCandidate(cert));
                }
            }
            finally { store.Close(); }
            return result;
        }

        private static CertificateCandidate ToCandidate(X509Certificate2 cert)
        {
            bool? keyEncipherment = null;
            List<string> ekus = null;
            var emails = new List<string>();

            string nameEmail = Safe(() => cert.GetNameInfo(X509NameType.EmailName, false), null);
            if (!string.IsNullOrWhiteSpace(nameEmail)) emails.Add(nameEmail);

            foreach (var ext in cert.Extensions)
            {
                var keyUsage = ext as X509KeyUsageExtension;
                if (keyUsage != null)
                {
                    keyEncipherment = (keyUsage.KeyUsages & X509KeyUsageFlags.KeyEncipherment) != 0;
                    continue;
                }
                var eku = ext as X509EnhancedKeyUsageExtension;
                if (eku != null)
                {
                    ekus = new List<string>();
                    foreach (var oid in eku.EnhancedKeyUsages) ekus.Add(oid.Value);
                    continue;
                }
                if (ext.Oid != null && ext.Oid.Value == OidSan)
                    emails.AddRange(EncryptionCertificatePolicy.ParseRfc822Names(ext.RawData));
            }

            string commonName = Safe(() => cert.GetNameInfo(X509NameType.SimpleName, false), cert.Subject);
            string display = emails.Count > 0 && !string.Equals(commonName, emails[0], StringComparison.OrdinalIgnoreCase)
                ? commonName + " <" + emails[0] + ">"
                : commonName;

            return new CertificateCandidate(cert.Thumbprint, display, cert.NotBefore, cert.NotAfter,
                                            Safe(() => cert.HasPrivateKey, false),
                                            cert.PublicKey?.Oid?.Value == OidRsa,
                                            keyEncipherment, ekus, emails.Distinct(StringComparer.OrdinalIgnoreCase).ToList());
        }

        /// <summary>The SMTP address of the account delivering to the default store, else the first account's.</summary>
        private string ReadAccountSmtp()
        {
            using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
            {
                string defaultStoreId = null;
                try
                {
                    using (var store = new ComRef<OutlookOM.Store>(session.Value.DefaultStore))
                        defaultStoreId = store.Value.StoreID;
                }
                catch { /* no default store: first account below */ }

                string first = null;
                using (var accounts = new ComRef<OutlookOM.Accounts>(session.Value.Accounts))
                {
                    int count = Safe(() => accounts.Value.Count, 0);
                    for (int i = 1; i <= count; i++)
                    {
                        using (var account = new ComRef<OutlookOM.Account>(accounts.Value[i]))
                        {
                            string smtp = Safe(() => account.Value.SmtpAddress, null);
                            if (first == null && !string.IsNullOrWhiteSpace(smtp)) first = smtp;
                            if (defaultStoreId == null) continue;

                            string deliveryStoreId = null;
                            try
                            {
                                using (var delivery = new ComRef<OutlookOM.Store>(account.Value.DeliveryStore))
                                    deliveryStoreId = delivery.Value.StoreID;
                            }
                            catch { /* some account types have no delivery store */ }
                            if (deliveryStoreId == defaultStoreId && !string.IsNullOrWhiteSpace(smtp)) return smtp;
                        }
                    }
                }
                return first;
            }
        }

        // ---- create & verify -------------------------------------------------------

        private sealed class Envelope
        {
            public string Subject;
            public bool UnRead;
            public OutlookOM.OlImportance Importance;
            public string Categories;
            public object[] Values;   // aligned with EnvelopeTags; null = absent on the source
            public readonly List<KeyValuePair<string, OutlookOM.OlMailRecipientType>> Recipients =
                new List<KeyValuePair<string, OutlookOM.OlMailRecipientType>>();
            public readonly List<string> RecipientNames = new List<string>();
        }

        public MailItemRef CreateEncryptedCopy(MailItemRef source, FolderNode destination, byte[] smimePayload, bool markAsRead)
        {
            if (source == null || destination == null || smimePayload == null) return null;

            string stageDir = null;
            string entryId = null;
            try
            {
                using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
                {
                    var envelope = ReadEnvelope(session.Value, source);
                    if (envelope == null) return null;

                    stageDir = Path.Combine(_tempDirectory, Guid.NewGuid().ToString("N"));
                    Directory.CreateDirectory(stageDir);
                    string payloadPath = Path.Combine(stageDir, "smime.p7m");  // ciphertext: safe on disk
                    File.WriteAllBytes(payloadPath, smimePayload);

                    Step("stub", source.EntryId);
                    entryId = CreateStub(session.Value, destination, envelope, payloadPath);
                    if (entryId == null) return null;

                    Step("reopen", entryId);
                    using (var comItem = new ComRef<object>(session.Value.GetItemFromID(entryId, destination.StoreId)))
                    {
                        var mail = comItem.Value as OutlookOM.MailItem;
                        if (mail == null)
                        {
                            Log.Warning("Re-encryption: the new item did not re-open as a mail item ({EntryId}).", entryId);
                            DeleteStub(session.Value, entryId, destination.StoreId);
                            return null;
                        }

                        // Checked while the item is still IPM.Note: on an S/MIME item,
                        // touching the attachment would trigger decryption.
                        Step("payload-check", entryId);
                        if (!PayloadMatches(mail, smimePayload))
                        {
                            Log.Warning("Re-encryption: the stored payload does not match ({EntryId}).", entryId);
                            DeleteStub(session.Value, entryId, destination.StoreId);
                            return null;
                        }

                        Step("envelope", entryId);
                        AddRecipients(mail, envelope);
                        ApplyEnvelope(mail, envelope);
                        mail.UnRead = !markAsRead && envelope.UnRead;
                        mail.Save();

                        Step("class-switch", entryId);
                        mail.MessageClass = "IPM.Note.SMIME";
                        mail.Save();
                        Step("saved", entryId);
                    }
                    Log.Debug("Re-encryption: encrypted copy created ({EntryId}).", entryId);
                    return new MailItemRef(destination.StoreId, entryId, source.Subject, false);
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Re-encryption: creating the encrypted copy failed ({EntryId}).", entryId ?? "(none)");
                if (entryId != null)
                {
                    try
                    {
                        using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
                            DeleteStub(session.Value, entryId, destination.StoreId);
                    }
                    catch { /* best effort */ }
                }
                return null;
            }
            finally
            {
                if (stageDir != null) ScrubDirectory(stageDir);
            }
        }

        /// <summary>
        /// Breadcrumb at Information (v2.14.0.0 B2): the target runs without
        /// diagnostic logging, and after a crash the last line names the step.
        /// </summary>
        private static void Step(string step, string entryId) =>
            Log.Information("Re-encryption step {Step} ({EntryId}).", step, entryId);

        /// <summary>The PostItem stage: subject, ciphertext attachment, then class IPM.Note. Returns the EntryID.</summary>
        private static string CreateStub(OutlookOM.NameSpace session, FolderNode destination, Envelope envelope, string payloadPath)
        {
            using (var folder = new ComRef<OutlookOM.Folder>((OutlookOM.Folder)session.GetFolderFromID(destination.EntryId, destination.StoreId)))
            using (var items = new ComRef<OutlookOM.Items>(folder.Value.Items))
            using (var post = new ComRef<OutlookOM.PostItem>((OutlookOM.PostItem)items.Value.Add(OutlookOM.OlItemType.olPostItem)))
            {
                post.Value.Subject = envelope.Subject ?? string.Empty;
                using (var attachments = new ComRef<OutlookOM.Attachments>(post.Value.Attachments))
                using (var att = new ComRef<OutlookOM.Attachment>(attachments.Value.Add(payloadPath, OutlookOM.OlAttachmentType.olByValue)))
                using (var pa = new ComRef<OutlookOM.PropertyAccessor>(att.Value.PropertyAccessor))
                    pa.Value.SetProperty(PR_ATTACH_MIME_TAG_W, "application/pkcs7-mime");

                // The delivery time may only stick before the first save; set again later.
                int delivery = Array.IndexOf(EnvelopeTags, PR_MESSAGE_DELIVERY_TIME);
                if (envelope.Values[delivery] != null)
                {
                    try
                    {
                        using (var pa = new ComRef<OutlookOM.PropertyAccessor>(post.Value.PropertyAccessor))
                            pa.Value.SetProperty(PR_MESSAGE_DELIVERY_TIME, envelope.Values[delivery]);
                    }
                    catch { /* retried after the class switch */ }
                }

                post.Value.Save();
                string entryId = post.Value.EntryID;
                post.Value.MessageClass = "IPM.Note";
                post.Value.Save();
                return entryId;
            }
        }

        private static Envelope ReadEnvelope(OutlookOM.NameSpace session, MailItemRef source)
        {
            using (var comItem = new ComRef<object>(session.GetItemFromID(source.EntryId, source.StoreId)))
            {
                var mail = comItem.Value as OutlookOM.MailItem;
                if (mail == null) return null;

                var envelope = new Envelope
                {
                    Subject = Safe(() => mail.Subject, string.Empty),
                    UnRead = Safe(() => mail.UnRead, false),
                    Importance = Safe(() => mail.Importance, OutlookOM.OlImportance.olImportanceNormal),
                    Categories = Safe(() => mail.Categories, null),
                };

                envelope.Values = new object[EnvelopeTags.Length];
                using (var pa = new ComRef<OutlookOM.PropertyAccessor>(mail.PropertyAccessor))
                {
                    for (int i = 0; i < EnvelopeTags.Length; i++)
                    {
                        // One by one: a missing property throws, and that is the normal case.
                        string tag = EnvelopeTags[i];
                        try { envelope.Values[i] = pa.Value.GetProperty(tag); }
                        catch { envelope.Values[i] = null; }
                    }
                }

                using (var recipients = new ComRef<OutlookOM.Recipients>(mail.Recipients))
                {
                    int count = Safe(() => recipients.Value.Count, 0);
                    int skipped = 0;
                    for (int i = 1; i <= count; i++)
                    {
                        using (var recipient = new ComRef<OutlookOM.Recipient>(recipients.Value[i]))
                        {
                            string smtp = null;
                            try
                            {
                                using (var rpa = new ComRef<OutlookOM.PropertyAccessor>(recipient.Value.PropertyAccessor))
                                    smtp = rpa.Value.GetProperty(PR_SMTP_ADDRESS_W) as string;
                            }
                            catch { /* not set on this recipient */ }
                            if (string.IsNullOrWhiteSpace(smtp))
                            {
                                string address = Safe(() => recipient.Value.Address, null);
                                if (address != null && address.Contains("@")) smtp = address;
                            }
                            if (string.IsNullOrWhiteSpace(smtp)) { skipped++; continue; }

                            var type = (OutlookOM.OlMailRecipientType)Safe(() => recipient.Value.Type, (int)OutlookOM.OlMailRecipientType.olTo);
                            envelope.Recipients.Add(new KeyValuePair<string, OutlookOM.OlMailRecipientType>(smtp, type));
                            envelope.RecipientNames.Add(Safe(() => recipient.Value.Name, smtp));
                        }
                    }
                    if (skipped > 0)
                        Log.Information("Re-encryption: {Skipped} recipient(s) without an SMTP address were not copied.", skipped);
                }
                return envelope;
            }
        }

        private static void AddRecipients(OutlookOM.MailItem mail, Envelope envelope)
        {
            using (var recipients = new ComRef<OutlookOM.Recipients>(mail.Recipients))
            {
                for (int i = 0; i < envelope.Recipients.Count; i++)
                {
                    using (var recipient = new ComRef<OutlookOM.Recipient>(recipients.Value.Add(envelope.Recipients[i].Key)))
                    {
                        recipient.Value.Type = (int)envelope.Recipients[i].Value;
                        try { recipient.Value.Resolve(); } catch { /* offline: stays a one-off */ }
                        try
                        {
                            using (var rpa = new ComRef<OutlookOM.PropertyAccessor>(recipient.Value.PropertyAccessor))
                                rpa.Value.SetProperty(PR_DISPLAY_NAME_W, envelope.RecipientNames[i]);
                        }
                        catch { /* keep the resolved or address display name */ }
                    }
                }
            }
        }

        private static void ApplyEnvelope(OutlookOM.MailItem mail, Envelope envelope)
        {
            using (var pa = new ComRef<OutlookOM.PropertyAccessor>(mail.PropertyAccessor))
            {
                var failed = new List<string>();
                for (int i = 0; i < EnvelopeTags.Length; i++)
                {
                    string tag = EnvelopeTags[i];
                    try
                    {
                        if (envelope.Values[i] != null) pa.Value.SetProperty(tag, envelope.Values[i]);
                        else if (i < 12) pa.Value.DeleteProperty(tag);   // identity tags only
                    }
                    catch { failed.Add(tag.Substring(Tag.Length)); }
                }
                if (failed.Count > 0)
                    Log.Debug("Re-encryption: envelope properties not set on the copy: {Tags}.", string.Join(",", failed));

                // The PostItem stage stamps the "post" icon (1), and deleting it does
                // not stick; -1 tells Outlook to derive the icon from the class.
                try { pa.Value.SetProperty(PR_ICON_INDEX, -1); } catch { /* cosmetic only */ }
            }

            try { mail.Importance = envelope.Importance; } catch { }
            if (!string.IsNullOrEmpty(envelope.Categories))
            {
                try { mail.Categories = envelope.Categories; } catch { }
            }
        }

        private bool PayloadMatches(OutlookOM.MailItem mail, byte[] payload)
        {
            using (var attachments = new ComRef<OutlookOM.Attachments>(mail.Attachments))
            {
                if (Safe(() => attachments.Value.Count, 0) != 1) return false;
                using (var att = new ComRef<OutlookOM.Attachment>(attachments.Value[1]))
                {
                    if (!ProtectedMailPolicy.IsSmimePayloadName(Safe(() => att.Value.FileName, null))) return false;
                    byte[] stored = ReadAttachmentBytes(att.Value);
                    return stored != null && stored.SequenceEqual(payload);
                }
            }
        }

        private static void DeleteStub(OutlookOM.NameSpace session, string entryId, string storeId)
        {
            try
            {
                using (var comItem = new ComRef<object>(session.GetItemFromID(entryId, storeId)))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    var post = comItem.Value as OutlookOM.PostItem;
                    if (mail != null) mail.Delete();
                    else post?.Delete();
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Re-encryption: deleting an incomplete copy failed ({EntryId}).", entryId);
            }
        }

        public bool VerifyEncryptedCopy(MailItemRef copy)
        {
            if (copy == null) return false;
            try
            {
                using (var session = new ComRef<OutlookOM.NameSpace>(_app.Session))
                using (var comItem = new ComRef<object>(session.Value.GetItemFromID(copy.EntryId, copy.StoreId)))
                {
                    var mail = comItem.Value as OutlookOM.MailItem;
                    if (mail == null) return false;
                    // The folder's table row only: never the body or attachments.
                    string stored = MailItemProbes.ReadStoredProtectedClass(mail);
                    bool ok = stored != null && stored.StartsWith("IPM.Note.SMIME", StringComparison.OrdinalIgnoreCase);
                    if (!ok) Log.Warning("Re-encryption: the copy is not stored as S/MIME ({EntryId}).", copy.EntryId);
                    return ok;
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Re-encryption: verifying the copy failed ({EntryId}).", copy.EntryId);
                return false;
            }
        }

        // ---- temp files -------------------------------------------------------------

        /// <summary>
        /// Zero then delete every file left in the temp directory (a crash between
        /// staging and cleanup). COM-free; safe to run off the UI thread at startup.
        /// </summary>
        public static void SweepTempDirectory(string directory)
        {
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory)) return;
            foreach (string file in Directory.GetFiles(directory, "*", SearchOption.AllDirectories))
                ScrubAndDelete(file);
            foreach (string dir in Directory.GetDirectories(directory))
            {
                try { Directory.Delete(dir, true); } catch { /* retried next start */ }
            }
        }

        private static void ScrubDirectory(string directory)
        {
            try
            {
                foreach (string file in Directory.GetFiles(directory)) ScrubAndDelete(file);
                Directory.Delete(directory, true);
            }
            catch { /* swept at next start */ }
        }

        /// <summary>
        /// Overwrite with zeros, flush to disk, then delete (3 tries: an antivirus may
        /// hold the file). Best-effort: SSD wear levelling or a scanner's copy can
        /// keep the bytes elsewhere, which the security documentation states.
        /// </summary>
        private static void ScrubAndDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    using (var stream = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough))
                    {
                        var zeros = new byte[64 * 1024];
                        long remaining = stream.Length;
                        while (remaining > 0)
                        {
                            int count = (int)Math.Min(zeros.Length, remaining);
                            stream.Write(zeros, 0, count);
                            remaining -= count;
                        }
                        stream.Flush(true);
                    }
                }
            }
            catch { /* still try to delete */ }

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    if (File.Exists(path)) File.Delete(path);
                    return;
                }
                catch { Thread.Sleep(50); }
            }
        }

        private static T Safe<T>(Func<T> read, T fallback)
        {
            try { return read(); }
            catch { return fallback; }
        }
    }
}
