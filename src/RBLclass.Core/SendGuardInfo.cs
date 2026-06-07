using System.Collections.Generic;

namespace RBLclass.Core
{
    /// <summary>
    /// What the send-time guards (legacy 6a-6b) need from a mail item that's
    /// about to go out, extracted by <see cref="IMailStore.InspectForSend"/>
    /// while the live COM item is still in hand on <c>Application.ItemSend</c>.
    /// Plain data so the guard decisions in <see cref="ForgottenAttachmentGuard"/>
    /// and <see cref="ExternalRecipientGuard"/> stay pure and testable.
    /// </summary>
    public sealed class SendGuardInfo
    {
        public SendGuardInfo(string body, int attachmentCount, IReadOnlyList<RecipientAddress> recipients)
        {
            Body = body;
            AttachmentCount = attachmentCount;
            Recipients = recipients;
        }

        public string Body { get; }

        public int AttachmentCount { get; }

        public IReadOnlyList<RecipientAddress> Recipients { get; }
    }
}
