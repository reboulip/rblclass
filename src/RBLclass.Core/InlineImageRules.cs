namespace RBLclass.Core
{
    /// <summary>
    /// Which attachments the disposition modal offers (v2.11.0.0 C2). Regular
    /// attachments always; body images (inline with a content id) only when the
    /// user opted in; OLE objects, embedded items and hidden parts never.
    /// </summary>
    public static class InlineImageRules
    {
        public static bool IsOffered(AttachmentInfo attachment, bool treatInlineImages) =>
            attachment != null
            && (!attachment.IsInline || (treatInlineImages && attachment.IsInlineImage));

        /// <summary>A content id as compared everywhere: trimmed, without the surrounding angle brackets.</summary>
        public static string NormalizeContentId(string contentId)
        {
            if (contentId == null) return null;
            string id = contentId.Trim();
            if (id.StartsWith("<") && id.EndsWith(">") && id.Length >= 2)
                id = id.Substring(1, id.Length - 2).Trim();
            return id.Length == 0 ? null : id;
        }
    }
}
