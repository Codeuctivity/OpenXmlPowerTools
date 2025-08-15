using Codeuctivity.OpenXmlPowerTools;
using SkiaSharp;
using Xunit;

namespace Codeuctivity.Tests
{
    public class ColorParserTests
    {
        [Theory]
        [InlineData("red", 255, 0, 0)]
        [InlineData("RED", 255, 0, 0)]
        [InlineData("#00FF00", 0, 255, 0)]
        [InlineData("#00ff00", 0, 255, 0)]
        [InlineData("blue", 0, 0, 255)]
        [InlineData("#0000FF", 0, 0, 255)]
        [InlineData("#abc", 170, 187, 204)]
        [InlineData("yellow", 255, 255, 0)]
        [InlineData("black", 0, 0, 0)]
        [InlineData("white", 255, 255, 255)]
        public void ShouldParseColors(string input, byte r, byte g, byte b)
        {
            var result = ColorParser.FromName(input);
            Assert.Equal(r, result.Red);
            Assert.Equal(g, result.Green);
            Assert.Equal(b, result.Blue);
        }

        [Theory]
        [InlineData("red", true)]
        [InlineData("#123456", true)]
        [InlineData("notacolor", false)]
        [InlineData("", false)]
        public void ShouldValidateColorNames(string input, bool valid)
        {
            Assert.Equal(valid, ColorParser.IsValidName(input));
        }

        [Fact]
        public void FromNameShouldThrowOnInvalid()
        {
            Assert.Throws<System.ArgumentException>(() => ColorParser.FromName("bogus"));
        }

        [Theory]
        [InlineData("notacolor")]
        [InlineData("#GGGGGG")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("   ")]
        public void ShouldRejectInvalidColors(string? input)
        {
            var success = ColorParser.TryFromName(input, out SKColor _);
            Assert.False(success);
        }
    }
}
