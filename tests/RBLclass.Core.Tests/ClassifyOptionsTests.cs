using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class ClassifyOptionsTests
    {
        [Fact]
        public void Default_instance_has_expected_boolean_defaults()
        {
            var o = new ClassifyOptions();

            o.KeepCopy.Should().BeFalse();
            o.RemoveAttachments.Should().BeFalse();
            o.MarkTasksComplete.Should().BeFalse();
            o.SafetyCopy.Should().BeFalse();
            o.MarkAsRead.Should().BeFalse();
            o.RecordHistory.Should().BeTrue(); // the one non-false default
            o.ReencryptProtected.Should().BeFalse();
            o.ReencryptCertificateThumbprint.Should().BeEmpty();
        }

        [Fact]
        public void Re_encryption_options_are_stored_and_passed_through_the_request()
        {
            var o = new ClassifyOptions(reencryptProtected: true, reencryptCertificateThumbprint: "AB12");
            var request = new ClassifyRequest(new MailItemRef[0], new FolderNode[0], o);

            request.ReencryptProtected.Should().BeTrue();
            request.ReencryptCertificateThumbprint.Should().Be("AB12");
        }

        [Fact]
        public void Default_static_field_matches_the_default_ctor()
        {
            var d = ClassifyOptions.Default;

            d.KeepCopy.Should().BeFalse();
            d.RemoveAttachments.Should().BeFalse();
            d.MarkTasksComplete.Should().BeFalse();
            d.SafetyCopy.Should().BeFalse();
            d.MarkAsRead.Should().BeFalse();
            d.RecordHistory.Should().BeTrue();
        }

        [Fact]
        public void Non_default_values_are_stored_and_returned()
        {
            var o = new ClassifyOptions(
                keepCopy: true, removeAttachments: true, markTasksComplete: true,
                safetyCopy: true, markAsRead: true, recordHistory: false);

            o.KeepCopy.Should().BeTrue();
            o.RemoveAttachments.Should().BeTrue();
            o.MarkTasksComplete.Should().BeTrue();
            o.SafetyCopy.Should().BeTrue();
            o.MarkAsRead.Should().BeTrue();
            o.RecordHistory.Should().BeFalse();
        }
    }
}
