using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    public class InlineImageRulesTests
    {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void A_regular_attachment_is_always_offered(bool option)
        {
            InlineImageRules.IsOffered(new AttachmentInfo(1, "a.pdf", 10), option).Should().BeTrue();
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void An_inline_part_without_a_content_id_is_never_offered(bool option)
        {
            InlineImageRules.IsOffered(new AttachmentInfo(1, "ole.bin", 10, isInline: true), option).Should().BeFalse();
        }

        [Fact]
        public void A_body_image_is_offered_only_when_opted_in()
        {
            var image = new AttachmentInfo(1, "image001.png", 10, isInline: true, contentId: "<image001@x>");

            image.IsInlineImage.Should().BeTrue();
            image.ContentId.Should().Be("image001@x");
            InlineImageRules.IsOffered(image, false).Should().BeFalse();
            InlineImageRules.IsOffered(image, true).Should().BeTrue();
        }

        [Fact]
        public void A_content_id_on_a_regular_attachment_is_ignored()
        {
            var file = new AttachmentInfo(1, "a.pdf", 10, isInline: false, contentId: "a@x");

            file.ContentId.Should().BeNull();
            file.IsInlineImage.Should().BeFalse();
        }

        [Theory]
        [InlineData(" <a@x> ", "a@x")]
        [InlineData("a@x", "a@x")]
        [InlineData("<>", null)]
        [InlineData("  ", null)]
        [InlineData(null, null)]
        public void Content_ids_are_normalized(string raw, string expected)
        {
            InlineImageRules.NormalizeContentId(raw).Should().Be(expected);
        }
    }
}
