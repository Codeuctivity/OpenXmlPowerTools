using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SkiaSharp;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace Codeuctivity.OpenXmlPowerTools
{
    /// <summary>
    /// Helper class for processing images in DocumentAssembler.
    /// Handles image insertion, dimension calculation, and format detection.
    /// </summary>
    internal static class ImageHelper
    {
        private const double EmusPerInch = 914400d;
        private const double EmusPerPoint = 12700d;
        private const double EmusPerPixel = 914400d / 96d;
        private const double EmusPerMillimeter = EmusPerInch / 25.4d;
        private const double EmusPerCentimeter = EmusPerMillimeter * 10d;

        /// <summary>
        /// Adds an image part to the appropriate OpenXml part type.
        /// </summary>
        internal static ImagePart AddImagePart(OpenXmlPart part)
        {
            return part switch
            {
                MainDocumentPart mainDocumentPart => mainDocumentPart.AddImagePart(ImagePartType.Png),
                HeaderPart headerPart => headerPart.AddImagePart(ImagePartType.Png),
                FooterPart footerPart => footerPart.AddImagePart(ImagePartType.Png),
                FootnotesPart footnotesPart => footnotesPart.AddImagePart(ImagePartType.Png),
                EndnotesPart endnotesPart => endnotesPart.AddImagePart(ImagePartType.Png),
                _ => throw new OpenXmlPowerToolsException($"Image: unsupported part type {part.GetType().Name}."),
            };
        }

        /// <summary>
        /// Creates an XML element representing an image in a Word document.
        /// </summary>
        internal static XElement CreateImageElement(string relationshipId, int docPrId, double widthEmu, double heightEmu, JustificationValues? justification)
        {
            var widthAttribute = widthEmu.ToString("0", CultureInfo.InvariantCulture);
            var heightAttribute = heightEmu.ToString("0", CultureInfo.InvariantCulture);
            XElement? paragraphProperties = null;
            if (justification.HasValue && justification.Value != JustificationValues.Left)
            {
                paragraphProperties = new XElement(W.pPr,
                    new XElement(W.jc, new XAttribute(W.val, ConvertJustificationToString(justification.Value))));
            }

            var pictureName = $"Picture {docPrId}";
            var element =
                new XElement(W.p,
                    paragraphProperties,
                    new XElement(W.r,
                        new XElement(W.drawing,
                            new XElement(WP.inline,
                                new XElement(WP.extent, new XAttribute("cx", widthAttribute), new XAttribute("cy", heightAttribute)),
                                new XElement(WP.effectExtent, new XAttribute("l", "0"), new XAttribute("t", "0"), new XAttribute("r", "0"), new XAttribute("b", "0")),
                                new XElement(WP.docPr, new XAttribute("id", docPrId), new XAttribute("name", pictureName)),
                                new XElement(WP.cNvGraphicFramePr,
                                    new XElement(A.graphicFrameLocks, new XAttribute("noChangeAspect", "1"))),
                                new XElement(A.graphic,
                                    new XElement(A.graphicData, new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
                                        new XElement(Pic._pic,
                                            new XElement(Pic.nvPicPr,
                                                new XElement(Pic.cNvPr, new XAttribute("id", "0"), new XAttribute("name", pictureName)),
                                                new XElement(Pic.cNvPicPr)),
                                            new XElement(Pic.blipFill,
                                                new XElement(A.blip, new XAttribute(R.embed, relationshipId)),
                                                new XElement(A.stretch,
                                                    new XElement(A.fillRect))),
                                            new XElement(Pic.spPr,
                                                new XElement(A.xfrm,
                                                    new XElement(A.off, new XAttribute("x", "0"), new XAttribute("y", "0")),
                                                    new XElement(A.ext, new XAttribute("cx", widthAttribute), new XAttribute("cy", heightAttribute))),
                                                new XElement(A.prstGeom, new XAttribute("prst", "rect"),
                                                    new XElement(A.avLst))))))))));
            return element;
        }

        /// <summary>
        /// Gets the next unique document property ID for an image.
        /// </summary>
        internal static int GetNextDocPrId(OpenXmlPart part)
        {
            var tracker = part.Annotation<ImageIdTracker>();
            if (tracker == null)
            {
                var existingIds = part
                    .GetXDocument()
                    .Descendants(WP.docPr)
                    .Select(dp => (int?)dp.Attribute("id") ?? 0);
                var maxId = existingIds.Any() ? existingIds.Max() : 0;
                tracker = new ImageIdTracker { NextId = maxId + 1 };
                part.AddAnnotation(tracker);
            }

            return tracker.NextId++;
        }

        /// <summary>
        /// Tracks the next available image ID for a document part.
        /// </summary>
        private sealed class ImageIdTracker
        {
            public int NextId { get; set; }
        }

        /// <summary>
        /// Tries to parse an alignment string into a JustificationValues enum.
        /// </summary>
        internal static bool TryGetJustification(string? align, out JustificationValues? justification, out string errorMessage)
        {
            justification = null;
            errorMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(align))
            {
                return true;
            }

            switch (align.Trim().ToLowerInvariant())
            {
                case "left":
                    justification = JustificationValues.Left;
                    return true;
                case "center":
                case "centre":
                    justification = JustificationValues.Center;
                    return true;
                case "right":
                    justification = JustificationValues.Right;
                    return true;
                case "justify":
                case "both":
                    justification = JustificationValues.Both;
                    return true;
                default:
                    errorMessage = "Image: Align attribute must be one of Left, Center, Right, or Justify.";
                    return false;
            }
        }

        /// <summary>
        /// Tries to parse a length string with units (px, pt, cm, mm, in, emu) into EMUs.
        /// </summary>
        internal static bool TryParseLengthToEmu(string? rawValue, out double? emuValue, out string errorMessage)
        {
            emuValue = null;
            errorMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return true;
            }

            var value = rawValue.Trim().ToLowerInvariant();
            double multiplier;
            if (value.EndsWith("px"))
            {
                multiplier = EmusPerPixel;
                value = value[..^2];
            }
            else if (value.EndsWith("pt"))
            {
                multiplier = EmusPerPoint;
                value = value[..^2];
            }
            else if (value.EndsWith("cm"))
            {
                multiplier = EmusPerCentimeter;
                value = value[..^2];
            }
            else if (value.EndsWith("mm"))
            {
                multiplier = EmusPerMillimeter;
                value = value[..^2];
            }
            else if (value.EndsWith("in"))
            {
                multiplier = EmusPerInch;
                value = value[..^2];
            }
            else if (value.EndsWith("emu"))
            {
                multiplier = 1d;
                value = value[..^3];
            }
            else
            {
                multiplier = EmusPerPixel;
            }

            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            {
                errorMessage = $"Image: Unable to parse length '{rawValue}'.";
                return false;
            }

            if (parsed <= 0)
            {
                errorMessage = $"Image: Length value '{rawValue}' must be greater than zero.";
                return false;
            }

            emuValue = parsed * multiplier;
            return true;
        }

        /// <summary>
        /// Converts a JustificationValues enum to its string representation.
        /// </summary>
        internal static string ConvertJustificationToString(JustificationValues value)
        {
            if (value == JustificationValues.Left)
            {
                return "left";
            }

            if (value == JustificationValues.Center)
            {
                return "center";
            }

            if (value == JustificationValues.Right)
            {
                return "right";
            }

            if (value == JustificationValues.Both)
            {
                return "both";
            }

            if (value == JustificationValues.Distribute)
            {
                return "distribute";
            }

            if (value == JustificationValues.Start)
            {
                return "start";
            }

            if (value == JustificationValues.End)
            {
                return "end";
            }

            return value.ToString().ToLowerInvariant();
        }

        /// <summary>
        /// Calculates image dimensions in EMUs, applying width, height, and max constraints while preserving aspect ratio.
        /// </summary>
        internal static bool TryCalculateImageDimensions(
            byte[] imageBytes,
            string? widthAttr,
            string? heightAttr,
            string? maxWidthAttr,
            string? maxHeightAttr,
            out double widthEmu,
            out double heightEmu,
            out string errorMessage)
        {
            widthEmu = 0;
            heightEmu = 0;
            errorMessage = string.Empty;

            if (!TryGetPixelSize(imageBytes, out var pixelWidth, out var pixelHeight, out errorMessage))
            {
                return false;
            }

            var actualWidthEmu = pixelWidth * EmusPerPixel;
            var actualHeightEmu = pixelHeight * EmusPerPixel;

            if (!TryParseLengthToEmu(widthAttr, out var widthOverrideEmu, out errorMessage))
            {
                return false;
            }

            if (!TryParseLengthToEmu(heightAttr, out var heightOverrideEmu, out errorMessage))
            {
                return false;
            }

            if (!TryParseLengthToEmu(maxWidthAttr, out var maxWidthEmu, out errorMessage))
            {
                return false;
            }

            if (!TryParseLengthToEmu(maxHeightAttr, out var maxHeightEmu, out errorMessage))
            {
                return false;
            }

            widthEmu = actualWidthEmu;
            heightEmu = actualHeightEmu;

            if (widthOverrideEmu.HasValue && heightOverrideEmu.HasValue)
            {
                widthEmu = widthOverrideEmu.Value;
                heightEmu = heightOverrideEmu.Value;
            }
            else if (widthOverrideEmu.HasValue)
            {
                widthEmu = widthOverrideEmu.Value;
                heightEmu = widthOverrideEmu.Value * actualHeightEmu / actualWidthEmu;
            }
            else if (heightOverrideEmu.HasValue)
            {
                heightEmu = heightOverrideEmu.Value;
                widthEmu = heightOverrideEmu.Value * actualWidthEmu / actualHeightEmu;
            }

            if (maxWidthEmu.HasValue && widthEmu > maxWidthEmu.Value)
            {
                var scale = maxWidthEmu.Value / widthEmu;
                widthEmu = maxWidthEmu.Value;
                heightEmu *= scale;
            }

            if (maxHeightEmu.HasValue && heightEmu > maxHeightEmu.Value)
            {
                var scale = maxHeightEmu.Value / heightEmu;
                heightEmu = maxHeightEmu.Value;
                widthEmu *= scale;
            }

            if (widthEmu <= 0 || heightEmu <= 0)
            {
                errorMessage = "Image: Calculated dimensions are invalid.";
                return false;
            }

            return true;
        }

        /// <summary>
        /// Gets pixel dimensions of an image, trying SkiaSharp first and falling back to header inspection.
        /// </summary>
        internal static bool TryGetPixelSize(byte[] imageBytes, out int width, out int height, out string errorMessage)
        {
            width = 0;
            height = 0;
            errorMessage = string.Empty;

            try
            {
                using var bitmap = SKBitmap.Decode(imageBytes);
                if (bitmap != null && bitmap.Width > 0 && bitmap.Height > 0)
                {
                    width = bitmap.Width;
                    height = bitmap.Height;
                    return true;
                }
            }
            catch
            {
                // ignore and fall back to header-based detection
            }

            if (TryReadPngDimensions(imageBytes, out width, out height))
            {
                return true;
            }

            if (TryReadJpegDimensions(imageBytes, out width, out height))
            {
                return true;
            }

            if (TryReadGifDimensions(imageBytes, out width, out height))
            {
                return true;
            }

            errorMessage = "Image: Unable to determine image dimensions.";
            return false;
        }

        /// <summary>
        /// Reads PNG image dimensions from the file header.
        /// </summary>
        private static bool TryReadPngDimensions(byte[] bytes, out int width, out int height)
        {
            width = 0;
            height = 0;
            if (bytes.Length < 24)
            {
                return false;
            }

            var signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
            for (var i = 0; i < signature.Length; i++)
            {
                if (bytes[i] != signature[i])
                {
                    return false;
                }
            }

            if (bytes[12] != 0x49 || bytes[13] != 0x48 || bytes[14] != 0x44 || bytes[15] != 0x52)
            {
                return false;
            }

            width = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
            height = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];
            return width > 0 && height > 0;
        }

        /// <summary>
        /// Reads JPEG image dimensions from the file header.
        /// </summary>
        private static bool TryReadJpegDimensions(byte[] bytes, out int width, out int height)
        {
            width = 0;
            height = 0;
            if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8)
            {
                return false;
            }

            var index = 2;
            while (index + 9 < bytes.Length)
            {
                if (bytes[index] != 0xFF)
                {
                    index++;
                    continue;
                }

                var marker = bytes[index + 1];
                index += 2;

                if (marker == 0xD8 || marker == 0xD9)
                {
                    continue;
                }

                if (index + 2 > bytes.Length)
                {
                    break;
                }

                var length = (bytes[index] << 8) + bytes[index + 1];
                if (length < 2 || index + length > bytes.Length)
                {
                    break;
                }

                if (marker >= 0xC0 && marker <= 0xC3)
                {
                    height = (bytes[index + 3] << 8) + bytes[index + 4];
                    width = (bytes[index + 5] << 8) + bytes[index + 6];
                    return width > 0 && height > 0;
                }

                index += length;
            }

            return false;
        }

        /// <summary>
        /// Reads GIF image dimensions from the file header.
        /// </summary>
        private static bool TryReadGifDimensions(byte[] bytes, out int width, out int height)
        {
            width = 0;
            height = 0;
            if (bytes.Length < 10 || bytes[0] != 'G' || bytes[1] != 'I' || bytes[2] != 'F')
            {
                return false;
            }

            width = bytes[6] | (bytes[7] << 8);
            height = bytes[8] | (bytes[9] << 8);
            return width > 0 && height > 0;
        }
    }
}
