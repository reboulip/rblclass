using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// Outlook-side operations of re-encrypted filing (v2.12.0.0, experimental):
    /// read a protected mail as decrypted in memory, encrypt rebuilt content as
    /// S/MIME, and create the filed item directly as <c>IPM.Note.SMIME</c>, so no
    /// plaintext is ever stored. Kept separate from <see cref="IMailStore"/> so
    /// the shipped store contract does not grow.
    /// </summary>
    /// <remarks>
    /// Contract: no member throws; a failure returns null/false after the
    /// implementation logged its reason (never a subject, file name or address).
    /// <see cref="Encrypt"/> and <see cref="ListEncryptionCertificates"/> are
    /// COM-free; every other member touches Outlook and must run on the STA thread.
    /// </remarks>
    public interface IMailRebuilder
    {
        /// <summary>
        /// The mail's content as decrypted in memory, or null when it is not
        /// readable decrypted (its S/MIME payload part still shows: the smart card
        /// is absent), cannot be rebuilt faithfully (embedded item, OLE or
        /// by-reference attachment), exceeds the size cap, or the read failed.
        /// Each <see cref="MailPart.AttachmentId"/> is the Outlook
        /// <c>Attachment.Index</c>; if any attachment cannot be read, the whole item
        /// is refused rather than returned without it.
        /// </summary>
        MailContent ReadDecryptedContent(MailItemRef item);

        /// <summary>
        /// The S/MIME (CMS EnvelopedData, AES-256-CBC) payload of
        /// <paramref name="mimeEntity"/> encrypted to the certificate, or null.
        /// </summary>
        byte[] Encrypt(byte[] mimeEntity, string certificateThumbprint);

        /// <summary>
        /// Create, in <paramref name="destination"/>, an <c>IPM.Note.SMIME</c> item
        /// whose only part is <paramref name="smimePayload"/> (<c>smime.p7m</c>),
        /// with the envelope (subject, sender, recipients, dates, message id,
        /// conversation, importance, categories, flag) copied from
        /// <paramref name="source"/>. The source is never saved.
        /// </summary>
        /// <param name="markAsRead">True: the copy is read; false: it keeps the source's read state.</param>
        /// <returns>The new item, or null (a partial item is deleted).</returns>
        MailItemRef CreateEncryptedCopy(MailItemRef source, FolderNode destination, byte[] smimePayload, bool markAsRead);

        /// <summary>
        /// True when the copy is stored as an S/MIME item. Reads the folder's
        /// contents table only: never the item's body or attachments, which would
        /// trigger decryption.
        /// </summary>
        bool VerifyEncryptedCopy(MailItemRef copy);

        /// <summary>
        /// The certificate to encrypt to: <paramref name="preferredThumbprint"/>
        /// when set and still a valid candidate (no silent switch when it is not),
        /// else Automatic. Null when nothing is usable.
        /// </summary>
        CertificateInfo ResolveCertificate(string preferredThumbprint);

        /// <summary>Every certificate that can be chosen in Settings.</summary>
        IReadOnlyList<CertificateInfo> ListEncryptionCertificates();
    }
}
