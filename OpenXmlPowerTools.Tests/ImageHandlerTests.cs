using System.IO;
using System.Xml.Linq;
using Codeuctivity.OpenXmlPowerTools;
using Codeuctivity.OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter;
using SkiaSharp;
using Xunit;

namespace Codeuctivity.Tests
{
    public class ImageHandlerTests
    {
        [Theory]
        [InlineData(SKEncodedImageFormat.Png, "image/png")]
        [InlineData(SKEncodedImageFormat.Jpeg, "image/jpeg")]
        [InlineData(SKEncodedImageFormat.Webp, "image/webp")]
        public void ShouldTransformImagesToDataUri(SKEncodedImageFormat format, string mime)
        {
            using var surface = SKSurface.Create(new SKImageInfo(10, 10));
            surface.Canvas.Clear(SKColors.Red);
            using var image = surface.Snapshot();
            using var ms = new MemoryStream();
            using (var data = image.Encode(format, 100))
            {
                data.SaveTo(ms);
            }
            ms.Position = 0;
            var info = new ImageInfo { Image = ms, AltText = "alt", ImgStyleAttribute = new XAttribute(NoNamespace.style, "width:10px") };
            var handler = new ImageHandler();
            var result = handler.TransformImage(info);
            var srcAttr = result.Attribute(NoNamespace.src);
            Assert.NotNull(srcAttr);
            Assert.StartsWith($"data:{mime};base64,", srcAttr.Value);
            Assert.Equal("width:10px", result.Attribute(NoNamespace.style)?.Value);
        }

        [Fact]
        public void ShouldTransformGifToDataUri()
        {
            var gif = System.Convert.FromBase64String("R0lGODlhAQABAPAAAP///wAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw==");
            using var ms = new MemoryStream(gif);
            var info = new ImageInfo { Image = ms };
            var handler = new ImageHandler();
            var result = handler.TransformImage(info);
            var srcAttr = result.Attribute(NoNamespace.src);
            Assert.NotNull(srcAttr);
            Assert.StartsWith("data:image/gif;base64,", srcAttr.Value);
        }

        [Fact]
        public void ShouldThrowOnInvalidImage()
        {
            using var ms = new MemoryStream(new byte[] { 1, 2, 3, 4 });
            var handler = new ImageHandler();
            Assert.ThrowsAny<System.Exception>(() => handler.TransformImage(new ImageInfo { Image = ms }));
        }

        [Fact]
        public void ShouldIncludeAltText()
        {
            using var surface = SKSurface.Create(new SKImageInfo(5, 5));
            surface.Canvas.Clear(SKColors.Blue);
            using var image = surface.Snapshot();
            using var ms = new MemoryStream();
            using (var data = image.Encode(SKEncodedImageFormat.Png, 100))
            {
                data.SaveTo(ms);
            }
            ms.Position = 0;
            var info = new ImageInfo { Image = ms, AltText = "demo" };
            var handler = new ImageHandler();
            var result = handler.TransformImage(info);
            Assert.Equal("demo", result.Attribute(NoNamespace.alt)?.Value);
        }

        [Fact]
        public void ShouldOmitAltTextWhenNotProvided()
        {
            using var surface = SKSurface.Create(new SKImageInfo(5, 5));
            surface.Canvas.Clear(SKColors.Blue);
            using var image = surface.Snapshot();
            using var ms = new MemoryStream();
            using (var data = image.Encode(SKEncodedImageFormat.Png, 100))
            {
                data.SaveTo(ms);
            }
            ms.Position = 0;
            var info = new ImageInfo { Image = ms };
            var handler = new ImageHandler();
            var result = handler.TransformImage(info);
            Assert.Null(result.Attribute(NoNamespace.alt));
        }
    }
}
