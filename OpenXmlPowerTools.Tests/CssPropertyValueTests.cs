using Codeuctivity.OpenXmlPowerTools;
using SkiaSharp;
using Xunit;

namespace Codeuctivity.Tests
{
    public class CssPropertyValueTests
    {
        [Fact]
        public void ShouldRecognizeNamedColor()
        {
            var value = new CssPropertyValue { Type = CssValueType.String, Value = "red" };
            Assert.True(value.IsColor);
            var color = value.ToColor();
            Assert.Equal(SKColors.Red, color);
        }

        [Fact]
        public void ShouldRecognizeHexColor()
        {
            var value = new CssPropertyValue { Type = CssValueType.Hex, Value = "#0000FF" };
            Assert.True(value.IsColor);
            var color = value.ToColor();
            Assert.Equal(SKColors.Blue, color);
        }

        [Fact]
        public void ShouldRejectNonColor()
        {
            var value = new CssPropertyValue { Type = CssValueType.String, Value = "1234" };
            Assert.False(value.IsColor);
        }

        [Fact]
        public void ShouldParseHexWithoutHash()
        {
            var value = new CssPropertyValue { Type = CssValueType.Hex, Value = "00FF00" };
            Assert.True(value.IsColor);
            var color = value.ToColor();
            Assert.Equal(0u, color.Red);
            Assert.Equal(255u, color.Green);
            Assert.Equal(0u, color.Blue);
        }
    }
}
