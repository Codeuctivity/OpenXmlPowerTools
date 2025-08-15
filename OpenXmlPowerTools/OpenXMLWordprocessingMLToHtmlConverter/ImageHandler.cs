using SkiaSharp;
using System;
using System.IO;
using System.Xml.Linq;

namespace Codeuctivity.OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter
{
    /// <summary>
    /// Default image handler
    /// </summary>
    public class ImageHandler : IImageHandler
    {
        /// <summary>
        /// Transforms OpenXml Images to HTML embeddable images
        /// </summary>
        /// <param name="imageInfo"></param>
        /// <returns></returns>
        public XElement TransformImage(ImageInfo imageInfo)
        {
            using var imageStream = new MemoryStream();
            imageInfo.Image.CopyTo(imageStream);
            var data = imageStream.ToArray();

            using var codec = SKCodec.Create(new SKMemoryStream(data));
            var mimeType = GetMimeType(codec.EncodedFormat);
            var base64 = Convert.ToBase64String(data);
            var imageSource = $"data:{mimeType};base64,{base64}";

            return new XElement(Xhtml.img, new XAttribute(NoNamespace.src, imageSource), imageInfo.ImgStyleAttribute, imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
        }

        private static string GetMimeType(SKEncodedImageFormat format) => format switch
        {
            SKEncodedImageFormat.Bmp => "image/bmp",
            SKEncodedImageFormat.Gif => "image/gif",
            SKEncodedImageFormat.Ico => "image/x-icon",
            SKEncodedImageFormat.Jpeg => "image/jpeg",
            SKEncodedImageFormat.Png => "image/png",
            SKEncodedImageFormat.Wbmp => "image/vnd.wap.wbmp",
            SKEncodedImageFormat.Webp => "image/webp",
            _ => "application/octet-stream",
        };
    }
}
