using Codeuctivity.OpenXmlPowerTools;
using SkiaSharp;
using System;
using System.IO;
using System.Xml.Linq;

namespace DocumentAssembler04
{
    internal class Program
    {
        private static void Main()
        {
            var outputDir = CreateOutputDirectory();

            var templatePath = LocateTemplate();
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"TemplateDocument.docx not found at {templatePath}");
                return;
            }

            var data = LoadOrCreateSampleData();

            var templateDocument = new WmlDocument(templatePath);
            var assembled = DocumentAssembler.AssembleDocument(templateDocument, data, out var templateError);
            if (templateError)
            {
                Console.WriteLine("The template produced validation errors. Inspect the generated document for highlighted issues.");
            }

            var outputDocx = Path.Combine(outputDir.FullName, "DocumentWithImage.docx");
            assembled.SaveAs(outputDocx);
            Console.WriteLine($"Generated document: {outputDocx}");
        }

        private static DirectoryInfo CreateOutputDirectory()
        {
            var now = DateTime.Now;
            var dirName = $"ExampleOutput-{now.Year - 2000:00}-{now.Month:00}-{now.Day:00}-{now.Hour:00}{now.Minute:00}{now.Second:00}";
            var directory = new DirectoryInfo(dirName);
            if (!directory.Exists)
            {
                directory.Create();
            }
            return directory;
        }

        private static string LocateTemplate()
        {
            var baseDir = AppContext.BaseDirectory;
            var templateInOutput = Path.Combine(baseDir, "TemplateDocument.docx");
            if (File.Exists(templateInOutput))
            {
                return templateInOutput;
            }

            // fallback to project directory (useful when running from source)
            return Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "TemplateDocument.docx");
        }

        private static XElement LoadOrCreateSampleData()
        {
            var samplePath = Path.Combine(AppContext.BaseDirectory, "SampleData.xml");
            if (File.Exists(samplePath))
            {
                return XElement.Load(samplePath);
            }

            var samples = new XElement("Images",
                new XElement("Image", GenerateImageBase64(400, 200, SKColors.SteelBlue, "IMG1")),
                new XElement("Image", GenerateImageBase64(800, 400, SKColors.Teal, "IMG2")),
                new XElement("Image", GenerateImageBase64(320, 640, SKColors.Purple, "IMG3")),
                new XElement("Image", GenerateImageBase64(600, 160, SKColors.IndianRed, "IMG4"))
            );
            return samples;
        }

        private static string GenerateImageBase64(int width, int height, SKColor background, string label)
        {
            using var surface = SKSurface.Create(new SKImageInfo(width, height));
            var canvas = surface.Canvas;
            canvas.Clear(background);

            using (var paint = new SKPaint())
            {
                paint.IsAntialias = true;
                paint.Shader = SKShader.CreateLinearGradient(
                    new SKPoint(0, 0),
                    new SKPoint(width, height),
                    new[] { background, background.WithAlpha(200), SKColors.White },
                    null,
                    SKShaderTileMode.Clamp);
                canvas.DrawRect(new SKRect(0, 0, width, height), paint);
            }

            using (var borderPaint = new SKPaint { Color = SKColors.White, StrokeWidth = Math.Max(width, height) / 60f, IsStroke = true, IsAntialias = true })
            {
                canvas.DrawRect(new SKRect(borderPaint.StrokeWidth, borderPaint.StrokeWidth, width - borderPaint.StrokeWidth, height - borderPaint.StrokeWidth), borderPaint);
            }

            using (var circlePaint = new SKPaint { Color = SKColors.OrangeRed, IsAntialias = true })
            {
                canvas.DrawCircle(width / 2f, height / 2f, Math.Min(width, height) / 4f, circlePaint);
            }

            using var font = new SKFont { Size = Math.Min(width, height) / 5f, Edging = SKFontEdging.Antialias };
            using (var textPaint = new SKPaint
            {
                Color = SKColors.White,
                IsAntialias = true
            })
            {
                canvas.DrawText(label, width / 2f, (height / 2f) + (font.Size / 3f), SKTextAlign.Center, font, textPaint);
            }

            canvas.Flush();

            using var image = surface.Snapshot();
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            return Convert.ToBase64String(data.ToArray());
        }
    }
}
