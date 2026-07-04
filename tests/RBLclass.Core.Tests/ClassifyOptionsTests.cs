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
