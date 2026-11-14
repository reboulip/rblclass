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
    /// <summary>Protected-mail guards and the diagnostic trace (v2.11.0.0 A3, #62).</summary>
    public class ClassifierServiceProtectionTests
    {
        private const int ArgPath = 1;
        private const int ArgClass = 3;
        private const int ArgFlags = 4;
        private const int ArgProtected = 5;
        private const int ArgStored = 6;
        private const int ArgReason = 9;

        private static MailItemRef Item(string entry) => new MailItemRef("s1", entry, "Secret subject");

        private static readonly FolderNode D1 = new FolderNode("s1", "d1", null, "d1", "Archive / A", isLeaf: true);
        private static readonly FolderNode D2 = new FolderNode("s1", "d2", null, "d2", "Archive / B", isLeaf: true);
        private static readonly FolderNode Trash = new FolderNode("s1", "trash", null, "trash", "Deleted", isLeaf: true);

        private static IMailStore Store()
        {
            var store = Substitute.For<IMailStore>();
            store.MoveItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>())
                 .Returns(ci => new MailItemRef("s1", "moved-" + ((MailItemRef)ci[0]).EntryId, "Secret subject"));
            store.CopyItemToFolder(Arg.Any<MailItemRef>(), Arg.Any<FolderNode>(), Arg.Any<bool>())
                 .Returns(ci => new MailItemRef("s1", "copy-" + ((FolderNode)ci[1]).EntryId, "Secret subject"));
            store.GetDeletedItemsFolder("s1").Returns(Trash);
            store.GetProtectionState(Arg.Any<MailItemRef>()).Returns(new MailProtectionState("IPM.Note", null));
            return store;
        }

        private static IAppLog DebugLog()
        {
            var log = Substitute.For<IAppLog>();
            log.IsDebugEnabled.Returns(true);
            return log;
        }

        private static ClassifyRequest Request(ClassifyOptions options, params FolderNode[] destinations) =>
            new ClassifyRequest(new[] { Item("e1") }, destinations, options);

        private static List<object[]> Traces(IAppLog log) =>
            log.ReceivedCalls()
               .Where(c => c.GetMethodInfo().Name == nameof(IAppLog.Debug)
                           && ((string)c.GetArguments()[0]).StartsWith("Protection trace"))
               .Select(c => (object[])c.GetArguments()[1])
               .ToList();

        private static IEnumerable<string> TracedPaths(IAppLog log) =>
            Traces(log).Select(a => (string)a[ArgPath]);

        [Fact]
        public void A_protected_flagged_item_is_filed_but_its_flag_is_left_alone()
        {
            var store = Store();
            store.IsFlaggedIncomplete(Arg.Any<MailItemRef>()).Returns(true);
            store.IsEncryptedMail(Arg.Any<MailItemRef>()).Returns(true);
            store.GetParentFolder(Arg.Any<MailItemRef>()).Returns(D2);

            var result = new ClassifierService(store).Classify(
                Request(new ClassifyOptions(markTasksComplete: true), D1));

            store.DidNotReceive().MarkTaskComplete(Arg.Any<MailItemRef>());
            result.Undo.Should().NotBeNull();
            result.Undo.CompletedFlags.Should().BeEmpty();
            result.OriginalsMoved.Should().Be(1);
            result.Errors.Should().Be(0);
        }

        [Fact]
        public void A_failing_protection_check_counts_as_protected()
        {
            var store = Store();
            store.IsFlaggedIncomplete(Arg.Any<MailItemRef>()).Returns(true);
            store.IsEncryptedMail(Arg.Any<MailItemRef>()).Returns(_ => throw new InvalidOperationException());

            var result = new ClassifierService(store).Classify(
                Request(new ClassifyOptions(markTasksComplete: true), D1));

            store.DidNotReceive().MarkTaskComplete(Arg.Any<MailItemRef>());
            result.OriginalsMoved.Should().Be(1);
        }

        [Fact]
        public void An_unprotected_flagged_item_still_gets_its_flag_completed()
        {
            var store = Store();
            store.IsFlaggedIncomplete(Arg.Any<MailItemRef>()).Returns(true);

            new ClassifierService(store).Classify(Request(new ClassifyOptions(markTasksComplete: true), D1));

            store.Received(1).MarkTaskComplete(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void Preflight_does_not_offer_a_protected_flagged_item()
        {
            var store = Store();
            var plain = Item("plain");
            var secure = Item("secure");
            store.IsFlaggedIncomplete(Arg.Any<MailItemRef>()).Returns(true);
            store.IsEncryptedMail(secure).Returns(true);

            var preflight = new ClassifierService(store).Preflight(new[] { plain, secure }, widenConversation: false);

            preflight.FlaggedIncomplete.Should().ContainSingle().Which.Should().Be(plain);
        }

        [Fact]
        public void Without_debug_logging_the_trace_reads_nothing()
        {
            var store = Store();
            var log = Substitute.For<IAppLog>();

            new ClassifierService(store, log: log).Classify(Request(ClassifyOptions.Default, D1));

            store.DidNotReceive().GetProtectionState(Arg.Any<MailItemRef>());
        }

        [Fact]
        public void A_move_traces_the_item_before_and_after()
        {
            var store = Store();
            var log = DebugLog();

            new ClassifierService(store, log: log).Classify(Request(ClassifyOptions.Default, D1));

            TracedPaths(log).Should().Equal("pending", "move");
        }

        [Fact]
        public void Extra_destinations_trace_as_copies()
        {
            var log = DebugLog();

            new ClassifierService(Store(), log: log).Classify(Request(ClassifyOptions.Default, D1, D2));

            TracedPaths(log).Should().Equal("pending", "copy", "move");
        }

        [Fact]
        public void A_safety_copy_is_traced_on_the_copy()
        {
            var log = DebugLog();

            new ClassifierService(Store(), log: log).Classify(
                Request(new ClassifyOptions(safetyCopy: true), D1));

            TracedPaths(log).Should().Equal("pending", "safety-copy", "move");
        }

        [Fact]
        public void Keep_a_copy_traces_the_kept_original()
        {
            var store = Store();
            var log = DebugLog();

            new ClassifierService(store, log: log).Classify(Request(new ClassifyOptions(keepCopy: true), D1));

            TracedPaths(log).Should().Equal("pending", "copy", "kept-original");
        }

        [Fact]
        public void The_trace_formats_flags_and_missing_values()
        {
            var store = Store();
            store.GetProtectionState(Arg.Any<MailItemRef>()).Returns(
                new MailProtectionState("IPM.Note.SMIME", 3),
                new MailProtectionState(null, null));
            var log = DebugLog();

            new ClassifierService(store, log: log).Classify(Request(ClassifyOptions.Default, D1));

            var traces = Traces(log);
            traces[0][ArgClass].Should().Be("IPM.Note.SMIME");
            traces[0][ArgFlags].Should().Be("0x00000003");
            traces[1][ArgClass].Should().Be("unreadable");
            traces[1][ArgFlags].Should().Be("absent");
        }

        [Fact]
        public void An_item_that_no_longer_resolves_traces_as_unresolved()
        {
            var store = Store();
            store.GetProtectionState(Arg.Any<MailItemRef>()).Returns((MailProtectionState)null);
            var log = DebugLog();

            new ClassifierService(store, log: log).Classify(Request(ClassifyOptions.Default, D1));

            Traces(log)[0][ArgClass].Should().Be("unresolved");
        }

        [Fact]
        public void A_failing_trace_never_fails_the_filing()
        {
            var store = Store();
            store.GetProtectionState(Arg.Any<MailItemRef>()).Returns(_ => throw new InvalidOperationException());

            var result = new ClassifierService(store, log: DebugLog()).Classify(Request(ClassifyOptions.Default, D1));

            result.ItemsProcessed.Should().Be(1);
            result.OriginalsMoved.Should().Be(1);
            result.Errors.Should().Be(0);
        }

        [Fact]
        public void Losing_protection_logs_a_warning()
        {
            var store = Store();
            store.GetProtectionState(Arg.Any<MailItemRef>()).Returns(
                new MailProtectionState("IPM.Note.SMIME", null),
                new MailProtectionState("IPM.Note", null));
            var log = DebugLog();

            new ClassifierService(store, log: log).Classify(Request(ClassifyOptions.Default, D1));

            log.Received(1).Warning(
                Arg.Is<string>(s => s.StartsWith("Protection trace")), Arg.Any<object[]>());
        }

        [Fact]
        public void Keeping_protection_logs_no_warning()
        {
            var store = Store();
            store.GetProtectionState(Arg.Any<MailItemRef>()).Returns(new MailProtectionState("IPM.Note.SMIME", null));
            var log = DebugLog();

            new ClassifierService(store, log: log).Classify(Request(ClassifyOptions.Default, D1));

            log.DidNotReceive().Warning(Arg.Any<string>(), Arg.Any<object[]>());
        }

        // #62: the trace must show a mail decrypted in memory (plain live class,
        // protected stored class) as protected, and say why.
        [Fact]
        public void The_trace_records_the_stored_class_and_the_reason()
        {
            var store = Store();
            store.GetProtectionState(Arg.Any<MailItemRef>()).Returns(
                new MailProtectionState("IPM.Note", 0, storedMessageClass: "IPM.Note.SMIME"));
            var log = DebugLog();

            new ClassifierService(store, log: log).Classify(Request(ClassifyOptions.Default, D1));

            var trace = Traces(log)[0];
            trace[ArgClass].Should().Be("IPM.Note");
            trace[ArgProtected].Should().Be(true);
            trace[ArgStored].Should().Be("IPM.Note.SMIME");
            trace[ArgReason].Should().Be(ProtectionReason.StoredMessageClass);
        }

        [Fact]
        public void A_store_reporting_protection_by_stored_class_keeps_attachments()
        {
            var store = Store();
            store.RemoveAttachments(Arg.Any<MailItemRef>()).Returns(false); // the adapter's guard fired

            var result = new ClassifierService(store).Classify(
                Request(new ClassifyOptions(removeAttachments: true), D1));

            result.EncryptedStripSkips.Should().Be(1);
            result.OriginalsMoved.Should().Be(1);
        }

        [Fact]
        public void The_trace_never_records_the_subject()
        {
            var log = DebugLog();

            new ClassifierService(Store(), log: log).Classify(Request(ClassifyOptions.Default, D1));

            Traces(log).SelectMany(a => a).Should().NotContain("Secret subject");
        }
    }
}
