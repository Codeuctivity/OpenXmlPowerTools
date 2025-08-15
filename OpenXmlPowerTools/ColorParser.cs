using SkiaSharp;
using System;
using System.Drawing;

namespace Codeuctivity.OpenXmlPowerTools
{
    public static class ColorParser
    {
        public static SKColor FromName(string name)
        {
            if (!TryFromName(name, out var color))
            {
                throw new ArgumentException("Invalid color name", nameof(name));
            }
            return color;
        }

        public static bool TryFromName(string? name, out SKColor color)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                color = default;
                return false;
            }

            try
            {
                var drawingColor = ColorTranslator.FromHtml(name);
                color = new SKColor(drawingColor.R, drawingColor.G, drawingColor.B, drawingColor.A);
                return true;
            }
            catch
            {
                color = default;
                return false;
            }
        }

        public static bool IsValidName(string name)
        {
            return TryFromName(name, out _);
        }
    }
}
