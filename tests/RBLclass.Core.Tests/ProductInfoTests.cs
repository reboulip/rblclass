using FluentAssertions;
using RBLclass.Core;
using Xunit;

namespace RBLclass.Core.Tests
{
    /// <summary>
    /// Step 0 smoke test: proves the test project references the core
    /// assembly, the xUnit + FluentAssertions stack builds and runs. Replaced
    /// by the folder-search corpus tests (the high-value surface) in Step 2.
    /// </summary>
    public class ProductInfoTests
    {
        [Fact]
        public void Name_is_RBLclass()
        {
            ProductInfo.Name.Should().Be("RBLclass");
        }
    }
}
