using System;
using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using NSubstitute;
using RBLclass.Core;
using RBLclass.Core.Logging;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>Experimental re-encrypted filing of protected mail (v2.12.0.0 A6, #62).</summary>
    public class ClassifierServiceReencryptTests
    {
        private const string Subject = "Secret subject";
        private const string Thumb = "44D69614E1D3FE6A33600009D821156AA16CB9BC";

        private static readonly FolderNode Inbox = new FolderNode("s1", "inbox", null, "Inbox", "Inbox", isLeaf: true);
        private static readonly FolderNode D1 = new FolderNode("s1", "d1", null, "d1", "Archive / A", isLeaf: true);
        private static readonly FolderNode D2 = new FolderNode("s1", "d2", null, "d2", "Archive / B", isLeaf: true);
        private static readonly FolderNode Trash = new FolderNode("s1", "trash", null, "trash", "Deleted", isLeaf: true);

        private static readonly CertificateInfo Cert = new CertificateInfo(Thumb, "RBLclass E2E Test", new DateTime(2028, 11, 14), true);

        private static readonly MailPart Report = new MailPart(1, "rapport.pdf", "application/pdf", new byte[] { 1, 2, 3 });
        private static readonly MailPart Logo = new MailPart(2, "logo.png", "image/png", new byte[] { 4 }, isInline: true, contentId: "logo@x");

        private static MailItemRef Item(string entry = "e1") => new MailItemRef("s1", entry, Subject);

        /// <summary>A protected source (encrypted stored class, decrypted in memory), plain filed/moved refs.</summary>
        private static IMailStore Store(string liveClass = "IPM.Note", string storedClass = "IPM.Note.SMIME",
                                        bool payloadVisible = false, bool trashCopyStillProtected = true)
        {
            var store = Substitute.For<IMailStore>();
            store.IsEncryptedMail(Arg.Any<MailItemRef>()).Returns(ci =>
            {
                var r = (MailItemRef)ci[0];
                if (r.EntryId.StartsWith("plain")) return false;
                if (r.EntryId.StartsWith("moved-")) return trashCopyStillProtected;
                return true;
            });
            store.GetProtectionState(Arg.Any<MailItemRef>())
                 .Returns(new MailProtectionState(liveClass, null, storedClass, payloadVisible));
            store.GetParentFolder(Arg.Any<MailItemRef>()).Returns(Inbox);
            store.GetDeletedItemsFolder("s1").Returns(Trash);
            store.GetConversationKey(Arg.Any<MailItemRef>()).Returns(ci => "conv-" + ((MailItemRef)ci[0]).EntryId);
            store.MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>())
                 .Returns(ci => new MailItemRef("s1", "moved-" + ((MailItemRef)ci[0]).EntryId, Subject));
            store.CopyItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>())
                 .Returns(ci => new MailItemRef("s1", "copy-" + ((FolderNode)ci[1]).EntryId, Subject));
            store.SaveAttachmentToFile(Arg.Any<MailItemRef>(), Arg.Any<int>(), Arg.Any<string>()).Returns(true);
            return store;
        }

        private sealed class Rebuilder
        {
            public readonly IMailRebuilder Mock = Substitute.For<IMailRebuilder>();
            public readonly List<byte[]> EncryptedEntities = new List<byte[]>();
            public readonly List<byte[]> EncryptInputs = new List<byte[]>();

            public Rebuilder(MailContent content = null, CertificateInfo cert = null, bool noCert = false)
            {
                Mock.ResolveCertificate(Arg.Any<string>()).Returns(noCert ? null : (cert ?? Cert));
                Mock.ReadDecryptedContent(Arg.Any<MailItemRef>())
                    .Returns(content ?? new MailContent("<html><body><p>Bonjour</p><img src=\"cid:logo@x\"></body></html>",
                                                        null, new[] { Report, Logo }));
                Mock.Encrypt(Arg.Any<byte[]>(), Arg.Any<string>()).Returns(ci =>
                {
                    var mime = (byte[])ci[0];
                    EncryptInputs.Add(mime);
                    EncryptedEntities.Add((byte[])mime.Clone());
                    return new byte[] { 0x30, 0x80 };
                });
                Mock.CreateEncryptedCopy(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<byte[]>(), Arg.Any<bool>())
                    .Returns(ci => new MailItemRef("s1", "smime-" + ((FolderNode)ci[1]).EntryId, Subject));
                Mock.VerifyEncryptedCopy(Arg.Any<MailItemRef>()).Returns(true);
            }

            public MimeNode LastEntity => MimeTestReader.Parse(EncryptedEntities.Last());
        }

        private static ClassifyOptions On(bool keepCopy = false, bool removeAttachments = false, bool markAsRead = false,
                                          string thumbprint = null) =>
            new ClassifyOptions(keepCopy: keepCopy, removeAttachments: removeAttachments, markAsRead: markAsRead,
                                reencryptProtected: true, reencryptCertificateThumbprint: thumbprint);

        /// <summary>
        /// Since v2.14.0.0 B3 re-encryption only runs when the request changes the
        /// content. This banner signature is in no test mail: the request asks for a
        /// strip (so re-encryption runs) without changing anything.
        /// </summary>
        private const string NoMatchBanner = "[[banner absent from every test mail]]";

        private static ClassifyRequest Request(ClassifyOptions options, FolderNode[] destinations,
                                               IReadOnlyList<AttachmentDisposition> rows = null,
                                               string banner = NoMatchBanner, params MailItemRef[] items) =>
            new ClassifyRequest(items.Length == 0 ? new[] { Item() } : items, destinations, options,
                                bannerSignature: banner, attachmentDispositions: rows);

        private static IEnumerable<string> FileNames(MimeNode node) =>
            new[] { node.FileName }.Where(n => n != null).Concat(node.Children.SelectMany(FileNames));

        // ---- only when the content changes (v2.14.0.0 B3) ---------------------------

        [Fact]
        public void A_plain_classify_of_a_protected_mail_never_opens_the_rebuilder()
        {
            var store = Store();
            var rb = new Rebuilder();

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(
                Request(On(), new[] { D1 }, banner: null));

            rb.Mock.ReceivedCalls().Should().BeEmpty();
            store.Received(1).MoveItemToFolder(Arg.Is<MailItemRef>(r => r.EntryId == "e1"), D1, false);
            result.Reencrypted.Should().Be(0);
            result.ReencryptFallbacks.Should().Be(0);
        }

        [Fact]
        public void Removing_attachments_re_encrypts()
        {
            var store = Store();
            var rb = new Rebuilder();

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(
                Request(On(removeAttachments: true), new[] { D1 }, banner: null));

            result.Reencrypted.Should().Be(1);
            rb.Mock.Received(1).CreateEncryptedCopy(Arg.Any<MailItemRef>(), D1, Arg.Any<byte[]>(), Arg.Any<bool>());
        }

        [Fact]
        public void A_disposition_row_that_deletes_an_attachment_re_encrypts()
        {
            var store = Store();
            var rb = new Rebuilder();
            var rows = new[] { new AttachmentDisposition(Item(), 1, "rapport.pdf", AttachmentDispositionAction.Delete) };

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(
                Request(On(), new[] { D1 }, rows: rows, banner: null));

            result.Reencrypted.Should().Be(1);
        }

        [Fact]
        public void Keep_only_disposition_rows_do_not_re_encrypt()
        {
            var store = Store();
            var rb = new Rebuilder();
            var rows = new[] { new AttachmentDisposition(Item(), 1, "rapport.pdf", AttachmentDispositionAction.Keep) };

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(
                Request(On(), new[] { D1 }, rows: rows, banner: null));

            result.Reencrypted.Should().Be(0);
            rb.Mock.DidNotReceive().ReadDecryptedContent(Arg.Any<MailItemRef>());
        }

        // ---- step breadcrumbs (v2.14.0.0 B2) ---------------------------------------

        [Fact]
        public void Every_step_leaves_an_information_breadcrumb_in_order()
        {
            var store = Store();
            var rb = new Rebuilder();
            var log = Substitute.For<IAppLog>();

            new ClassifierService(store, log: log, rebuilder: rb.Mock).Classify(Request(On(), new[] { D1, D2 }));

            var steps = log.ReceivedCalls()
                .Where(c => c.GetMethodInfo().Name == nameof(IAppLog.Information)
                            && (string)c.GetArguments()[0] == "Re-encrypted filing step {Step} ({EntryId})")
                .Select(c => (string)((object[])c.GetArguments()[1])[0])
                .ToList();
            steps.Should().Equal("eligibility", "certificate", "read", "build", "encrypt",
                                 "create", "verify", "create", "verify", "original");
        }

        // ---- happy path ---------------------------------------------------------

        [Fact]
        public void An_eligible_mail_is_filed_as_an_encrypted_copy_and_the_original_goes_to_deleted_items()
        {
            var store = Store();
            var rb = new Rebuilder();

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(), new[] { D1 }));

            rb.Mock.Received(1).CreateEncryptedCopy(Arg.Is<MailItemRef>(r => r.EntryId == "e1"), D1, Arg.Any<byte[]>(), false);
            rb.Mock.Received(1).VerifyEncryptedCopy(Arg.Is<MailItemRef>(r => r.EntryId == "smime-d1"));
            store.Received(1).MoveItemToFolder(Arg.Is<MailItemRef>(r => r.EntryId == "e1"), Trash, false);
            store.DidNotReceive().MoveItemToFolder(Arg.Any<MailItemRef>(), D1, Arg.Any<bool>());
            store.DidNotReceive().CopyItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>());
            store.DidNotReceive().DeleteItem(Arg.Any<MailItemRef>());

            result.ItemsProcessed.Should().Be(1);
            result.Reencrypted.Should().Be(1);
            result.ReencryptFallbacks.Should().Be(0);
            result.CopiesMade.Should().Be(0);
            result.OriginalsMoved.Should().Be(0);
            result.Errors.Should().Be(0);
            result.Undo.Should().BeNull();
        }

        [Fact]
        public void The_certificate_chosen_in_settings_is_used()
        {
            var store = Store();
            var rb = new Rebuilder();

            new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(thumbprint: Thumb), new[] { D1 }));

            rb.Mock.Received(1).ResolveCertificate(Thumb);
            rb.Mock.Received(1).Encrypt(Arg.Any<byte[]>(), Thumb);
        }

        [Fact]
        public void Every_destination_gets_its_own_encrypted_copy()
        {
            var rb = new Rebuilder();

            var result = new ClassifierService(Store(), rebuilder: rb.Mock).Classify(Request(On(), new[] { D1, D2 }));

            rb.Mock.Received(1).CreateEncryptedCopy(Arg.Any<MailItemRef>(), D1, Arg.Any<byte[]>(), Arg.Any<bool>());
            rb.Mock.Received(1).CreateEncryptedCopy(Arg.Any<MailItemRef>(), D2, Arg.Any<byte[]>(), Arg.Any<bool>());
            rb.Mock.Received(1).Encrypt(Arg.Any<byte[]>(), Arg.Any<string>()); // built and encrypted once
            result.Reencrypted.Should().Be(1);
        }

        [Fact]
        public void Mark_as_read_is_passed_to_the_copy()
        {
            var rb = new Rebuilder();

            new ClassifierService(Store(), rebuilder: rb.Mock).Classify(Request(On(markAsRead: true), new[] { D1 }));

            rb.Mock.Received(1).CreateEncryptedCopy(Arg.Any<MailItemRef>(), D1, Arg.Any<byte[]>(), true);
        }

        [Fact]
        public void Keep_a_copy_leaves_the_original_in_place()
        {
            var store = Store();
            var rb = new Rebuilder();

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(keepCopy: true), new[] { D1 }));

            store.DidNotReceive().MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>());
            result.Reencrypted.Should().Be(1);
        }

        [Fact]
        public void An_original_written_decrypted_into_deleted_items_is_deleted_for_good()
        {
            var store = Store(trashCopyStillProtected: false);
            var log = Substitute.For<IAppLog>();

            new ClassifierService(store, log: log, rebuilder: new Rebuilder().Mock).Classify(Request(On(), new[] { D1 }));

            store.Received(1).DeleteItem(Arg.Is<MailItemRef>(r => r.EntryId == "moved-e1"));
            store.DidNotReceive().DeleteItemPermanently(Arg.Any<MailItemRef>());
            log.Received().Warning(Arg.Is<string>(t => t.Contains("written decrypted")), Arg.Any<object[]>());
        }

        [Fact]
        public void An_original_already_in_deleted_items_is_not_moved()
        {
            var store = Store();
            store.GetParentFolder(Arg.Any<MailItemRef>()).Returns(Trash);

            var result = new ClassifierService(store, rebuilder: new Rebuilder().Mock).Classify(Request(On(), new[] { D1 }));

            store.DidNotReceive().MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>());
            result.Errors.Should().Be(0);
        }

        [Fact]
        public void A_failed_move_to_deleted_items_keeps_the_filing_and_counts_an_error()
        {
            var store = Store();
            store.MoveItemToFolder(Arg.Any<MailItemRef>(), Trash, false).Returns((MailItemRef)null);

            var result = new ClassifierService(store, rebuilder: new Rebuilder().Mock).Classify(Request(On(), new[] { D1 }));

            result.Reencrypted.Should().Be(1);
            result.Errors.Should().Be(1);
        }

        // ---- content changes ----------------------------------------------------

        [Fact]
        public void Remove_attachments_strips_the_rebuilt_content_and_keeps_body_images()
        {
            var rb = new Rebuilder();

            new ClassifierService(Store(), rebuilder: rb.Mock).Classify(Request(On(removeAttachments: true), new[] { D1 }));

            FileNames(rb.LastEntity).Should().Equal("logo.png");
        }

        [Fact]
        public void Without_remove_attachments_every_part_is_kept()
        {
            var rb = new Rebuilder();

            new ClassifierService(Store(), rebuilder: rb.Mock).Classify(Request(On(), new[] { D1 }));

            FileNames(rb.LastEntity).Should().BeEquivalentTo("rapport.pdf", "logo.png");
        }

        [Fact]
        public void Save_to_rows_are_saved_from_the_source_then_removed()
        {
            var store = Store();
            var rb = new Rebuilder();
            var rows = new[] { new AttachmentDisposition(Item(), 1, "rapport.pdf", AttachmentDispositionAction.SaveTo, @"C:\Archive") };

            new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(removeAttachments: true), new[] { D1 }, rows));

            store.Received(1).SaveAttachmentToFile(Arg.Is<MailItemRef>(r => r.EntryId == "e1"), 1, @"C:\Archive");
            FileNames(rb.LastEntity).Should().Equal("logo.png");
        }

        [Fact]
        public void A_failed_save_keeps_every_attachment_and_counts_an_error()
        {
            var store = Store();
            store.SaveAttachmentToFile(Arg.Any<MailItemRef>(), Arg.Any<int>(), Arg.Any<string>()).Returns(false);
            var rb = new Rebuilder();
            var rows = new[] { new AttachmentDisposition(Item(), 1, "rapport.pdf", AttachmentDispositionAction.SaveTo, @"C:\Archive") };

            var result = new ClassifierService(store, rebuilder: rb.Mock)
                .Classify(Request(On(removeAttachments: true), new[] { D1 }, rows));

            FileNames(rb.LastEntity).Should().BeEquivalentTo("rapport.pdf", "logo.png");
            result.Errors.Should().Be(1);
            result.Reencrypted.Should().Be(1);
        }

        [Fact]
        public void The_banner_is_stripped_from_the_rebuilt_body()
        {
            const string banner = "<table><tr><td><b>CAUTION:</b> This email originated from outside the organisation. " +
                                  "Do not click links or open attachments unless you recognise the sender.</td></tr></table>";
            var rb = new Rebuilder(new MailContent("<html><body>" + banner + "<p>Kept text</p></body></html>", null, new MailPart[0]));

            new ClassifierService(Store(), rebuilder: rb.Mock).Classify(Request(On(), new[] { D1 }, banner: banner));

            string body = rb.LastEntity.Text;
            body.Should().Contain("Kept text").And.NotContain("CAUTION");
        }

        [Fact]
        public void The_plaintext_entity_is_wiped_after_encryption()
        {
            var rb = new Rebuilder();

            new ClassifierService(Store(), rebuilder: rb.Mock).Classify(Request(On(), new[] { D1 }));

            rb.EncryptInputs.Single().Should().OnlyContain(b => b == 0);
        }

        // ---- fallbacks ----------------------------------------------------------

        private static void ShouldHaveTakenTheStandardPath(IMailStore store, ClassifyResult result)
        {
            store.Received(1).MoveItemToFolder(Arg.Is<MailItemRef>(r => r.EntryId == "e1"), D1, false);
            store.DidNotReceive().MoveItemToFolder(Arg.Any<MailItemRef>(), Trash, Arg.Any<bool>());
            result.ItemsProcessed.Should().Be(1);
            result.OriginalsMoved.Should().Be(1);
            result.Reencrypted.Should().Be(0);
        }

        [Fact]
        public void No_certificate_falls_back_and_is_reported_once_per_classifier()
        {
            var store = Store();
            var rb = new Rebuilder(noCert: true);
            var service = new ClassifierService(store, rebuilder: rb.Mock);

            var first = service.Classify(Request(On(), new[] { D1 }));
            var second = service.Classify(Request(On(), new[] { D1 }, items: Item("e2")));

            store.Received(1).MoveItemToFolder(Arg.Is<MailItemRef>(r => r.EntryId == "e1"), D1, false);
            first.ReencryptFallbacks.Should().Be(1);
            first.ReencryptNoCertificate.Should().BeTrue();
            second.ReencryptFallbacks.Should().Be(1);
            second.ReencryptNoCertificate.Should().BeFalse();
            rb.Mock.DidNotReceive().ReadDecryptedContent(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void Unreadable_content_falls_back()
        {
            var store = Store();
            var rb = new Rebuilder();
            rb.Mock.ReadDecryptedContent(Arg.Any<MailItemRef>()).Returns((MailContent)null);

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(), new[] { D1 }));

            ShouldHaveTakenTheStandardPath(store, result);
            result.ReencryptFallbacks.Should().Be(1);
        }

        [Fact]
        public void An_attachment_map_mismatch_falls_back()
        {
            var store = Store();
            var rb = new Rebuilder();
            var rows = new[]
            {
                new AttachmentDisposition(Item(), 1, "renamed.pdf", AttachmentDispositionAction.Delete),
                new AttachmentDisposition(Item(), 7, "other.txt", AttachmentDispositionAction.Keep),
            };

            var result = new ClassifierService(store, rebuilder: rb.Mock)
                .Classify(Request(On(removeAttachments: true), new[] { D1 }, rows));

            ShouldHaveTakenTheStandardPath(store, result);
            rb.Mock.DidNotReceive().Encrypt(Arg.Any<byte[]>(), Arg.Any<string>());
        }

        [Fact]
        public void A_failed_encryption_falls_back()
        {
            var store = Store();
            var rb = new Rebuilder();
            rb.Mock.Encrypt(Arg.Any<byte[]>(), Arg.Any<string>()).Returns((byte[])null);

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(), new[] { D1 }));

            ShouldHaveTakenTheStandardPath(store, result);
            rb.Mock.DidNotReceive().CreateEncryptedCopy(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<byte[]>(), Arg.Any<bool>());
        }

        [Fact]
        public void A_failed_copy_removes_the_copies_already_made_and_falls_back()
        {
            var store = Store();
            var rb = new Rebuilder();
            rb.Mock.CreateEncryptedCopy(Arg.Any<MailItemRef>(), D2, Arg.Any<byte[]>(), Arg.Any<bool>()).Returns((MailItemRef)null);

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(), new[] { D1, D2 }));

            store.Received(1).DeleteItemPermanently(Arg.Is<MailItemRef>(r => r.EntryId == "smime-d1"));
            store.Received(1).CopyItemToFolder(Arg.Is<MailItemRef>(r => r.EntryId == "e1"), D1, false);
            store.Received(1).MoveItemToFolder(Arg.Is<MailItemRef>(r => r.EntryId == "e1"), D2, false);
            result.ReencryptFallbacks.Should().Be(1);
            result.Reencrypted.Should().Be(0);
        }

        [Fact]
        public void A_copy_that_does_not_verify_is_removed_and_falls_back()
        {
            var store = Store();
            var rb = new Rebuilder();
            rb.Mock.VerifyEncryptedCopy(Arg.Any<MailItemRef>()).Returns(false);

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(), new[] { D1 }));

            store.Received(1).DeleteItemPermanently(Arg.Is<MailItemRef>(r => r.EntryId == "smime-d1"));
            ShouldHaveTakenTheStandardPath(store, result);
            result.ReencryptFallbacks.Should().Be(1);
        }

        [Fact]
        public void A_throwing_rebuilder_falls_back()
        {
            var store = Store();
            var rb = new Rebuilder();
            rb.Mock.ReadDecryptedContent(Arg.Any<MailItemRef>()).Returns(_ => throw new OutOfMemoryException());

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(), new[] { D1 }));

            ShouldHaveTakenTheStandardPath(store, result);
        }

        // ---- not applicable: standard path, not counted -------------------------

        [Theory]
        [InlineData("IPM.Note.SMIME.MultipartSigned", "IPM.Note.SMIME.MultipartSigned", false)] // signed only
        [InlineData("IPM.Note", "IPM.Note.SMIME.MultipartSigned", false)]                      // signed only, stored
        [InlineData("IPM.Note.SMIME", "IPM.Note.SMIME", true)]                                 // card absent
        [InlineData("IPM.Note", "IPM.Note", false)]                                            // protected by unsaved changes only
        public void An_ineligible_protected_mail_takes_the_standard_path_silently(string live, string stored, bool payload)
        {
            var store = Store(live, stored, payload);
            var rb = new Rebuilder();

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(), new[] { D1 }));

            ShouldHaveTakenTheStandardPath(store, result);
            result.ReencryptFallbacks.Should().Be(0);
            rb.Mock.DidNotReceive().ResolveCertificate(Arg.Any<string>());
        }

        [Fact]
        public void A_mail_still_showing_its_payload_part_takes_the_standard_path_silently()
        {
            // Live class protected: the store does not probe the payload, the parts show it.
            var store = Store("IPM.Note.SMIME", "IPM.Note.SMIME");
            store.GetAttachments(Arg.Any<MailItemRef>()).Returns(new[] { new AttachmentInfo(1, "smime.p7m", 100) });
            var rb = new Rebuilder();

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(), new[] { D1 }));

            ShouldHaveTakenTheStandardPath(store, result);
            result.ReencryptFallbacks.Should().Be(0);
            rb.Mock.ReceivedCalls().Should().BeEmpty();
        }

        [Fact]
        public void A_decrypted_in_memory_mail_is_not_asked_for_its_parts()
        {
            var store = Store();

            new ClassifierService(store, rebuilder: new Rebuilder().Mock).Classify(Request(On(), new[] { D1 }));

            store.DidNotReceive().GetAttachments(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void A_plain_mail_takes_the_standard_path()
        {
            var store = Store();
            var rb = new Rebuilder();
            var plain = new MailItemRef("s1", "plain1", Subject);

            var result = new ClassifierService(store, rebuilder: rb.Mock).Classify(Request(On(), new[] { D1 }, items: plain));

            store.Received(1).MoveItemToFolder(plain, D1, false);
            rb.Mock.ReceivedCalls().Should().BeEmpty();
            result.Reencrypted.Should().Be(0);
        }

        [Fact]
        public void With_the_setting_off_the_standard_path_is_unchanged()
        {
            var store = Store();
            var rb = new Rebuilder();

            var result = new ClassifierService(store, rebuilder: rb.Mock)
                .Classify(Request(new ClassifyOptions(), new[] { D1 }));

            ShouldHaveTakenTheStandardPath(store, result);
            rb.Mock.ReceivedCalls().Should().BeEmpty();
            store.DidNotReceive().GetProtectionState(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void Without_a_rebuilder_the_setting_changes_nothing()
        {
            var store = Store();

            var result = new ClassifierService(store).Classify(Request(On(), new[] { D1 }));

            ShouldHaveTakenTheStandardPath(store, result);
            result.ReencryptFallbacks.Should().Be(0);
        }

        // ---- batches, history, logging ------------------------------------------

        [Fact]
        public void In_a_mixed_batch_only_the_plain_mail_is_undoable()
        {
            var store = Store();
            var history = Substitute.For<IClassificationHistory>();
            var plain = new MailItemRef("s1", "plain1", Subject);

            var result = new ClassifierService(store, history, rebuilder: new Rebuilder().Mock)
                .Classify(Request(On(), new[] { D1 }, items: new[] { Item(), plain }));

            result.ItemsProcessed.Should().Be(2);
            result.Reencrypted.Should().Be(1);
            result.Undo.Should().NotBeNull();
            result.Undo.Moves.Should().ContainSingle().Which.Current.EntryId.Should().Be("moved-plain1");
            result.Undo.HistoryBatchIds.Should().ContainSingle();

            string undoBatch = result.Undo.HistoryBatchIds[0];
            history.Received(1).Record(undoBatch, "conv-plain1", Arg.Any<IEnumerable<FolderNode>>(), Arg.Any<DateTime>());
            history.Received(1).Record(Arg.Is<string>(b => b != undoBatch), "conv-e1",
                                       Arg.Any<IEnumerable<FolderNode>>(), Arg.Any<DateTime>());
        }

        [Fact]
        public void A_re_encrypted_filing_still_feeds_auto_class_without_an_undo()
        {
            var history = Substitute.For<IClassificationHistory>();

            var result = new ClassifierService(Store(), history, rebuilder: new Rebuilder().Mock)
                .Classify(Request(On(), new[] { D1 }));

            history.Received(1).Record(Arg.Any<string>(), "conv-e1", Arg.Any<IEnumerable<FolderNode>>(), Arg.Any<DateTime>());
            result.Undo.Should().BeNull();
        }

        [Fact]
        public void The_certificate_is_resolved_once_per_classify()
        {
            var rb = new Rebuilder();

            new ClassifierService(Store(), rebuilder: rb.Mock)
                .Classify(Request(On(), new[] { D1 }, items: new[] { Item("e1"), Item("e2") }));

            rb.Mock.Received(1).ResolveCertificate(Arg.Any<string>());
            rb.Mock.Received(2).Encrypt(Arg.Any<byte[]>(), Arg.Any<string>());
        }

        [Fact]
        public void The_subject_is_never_logged()
        {
            var store = Store(trashCopyStillProtected: false);
            var rb = new Rebuilder();
            rb.Mock.VerifyEncryptedCopy(Arg.Is<MailItemRef>(r => r.EntryId == "smime-d2")).Returns(false);
            var log = Substitute.For<IAppLog>();
            log.IsDebugEnabled.Returns(true);
            var service = new ClassifierService(store, log: log, rebuilder: rb.Mock);

            service.Classify(Request(On(), new[] { D1 }));
            service.Classify(Request(On(), new[] { D1, D2 }, items: Item("e2")));

            var logged = log.ReceivedCalls()
                .SelectMany(c => c.GetArguments())
                .SelectMany(a => a is object[] arr ? arr : new[] { a })
                .Select(a => a?.ToString() ?? string.Empty);
            logged.Should().NotContain(a => a.Contains(Subject));
            log.Received().Information(Arg.Is<string>(t => t.StartsWith("Re-encrypted filing:")), Arg.Any<object[]>());
            log.Received().Warning(Arg.Is<string>(t => t.StartsWith("Re-encrypted filing fell back")), Arg.Any<object[]>());
        }

        [Fact]
        public void Auto_class_sums_the_re_encryption_counts()
        {
            var store = Store();
            var history = Substitute.For<IClassificationHistory>();
            history.GetLatestDestinations(Arg.Any<string>(), Arg.Any<DateTime>())
                   .Returns(new[] { new HistoryDestination("s1", "d1") });
            var rb = new Rebuilder();
            rb.Mock.ReadDecryptedContent(Arg.Is<MailItemRef>(r => r.EntryId == "e2")).Returns((MailContent)null);

            var result = new ClassifierService(store, history, rebuilder: rb.Mock).AutoClassify(
                new[] { Item("e1"), Item("e2") }, (s, e) => D1, On(), bannerSignature: NoMatchBanner);

            result.Filed.Should().Be(2);
            result.Reencrypted.Should().Be(1);
            result.ReencryptFallbacks.Should().Be(1);
            rb.Mock.Received(1).ResolveCertificate(Arg.Any<string>());
        }
    }
}
