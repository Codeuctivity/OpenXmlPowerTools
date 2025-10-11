using Codeuctivity.HtmlRenderer;
using Codeuctivity.OpenXmlPowerTools.OpenXMLWordprocessingMLToHtmlConverter;
using Codeuctivity.SkiaSharpCompare;
using DocumentFormat.OpenXml.Packaging;
using PuppeteerSharp;
using SkiaSharp;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xunit;

namespace Codeuctivity.Tests.OpenXMLWordProcessingMLToHtmlConverter
{
    public class WmlToHtmlConverterTests
    {
        // PowerShell one liner that generates InlineData for all files in a directory
        // dir | % { '[InlineData("' + $_.Name + '")]' } | clip

        [Theory]
        [InlineData("HC001-5DayTourPlanTemplate.docx", 0, false)]
        [InlineData("HC002-Hebrew-01.docx", 0, true)]
        [InlineData("HC003-Hebrew-02.docx", 0, true)]
        [InlineData("HC004-ResumeTemplate.docx", 0, false)]
        [InlineData("HC005-TaskPlanTemplate.docx", 0, false)]
        [InlineData("HC006-Test-01.docx", 0, true)]
        [InlineData("HC007-Test-02.docx", 0, true)]
        [InlineData("HC008-Test-03.docx", 0, false)]
        [InlineData("HC009-Test-04.docx", 0, true)]
        [InlineData("HC010-Test-05.docx", 0, true)]
        [InlineData("HC011-Test-06.docx", 0, false)]
        [InlineData("HC012-Test-07.docx", 0, true)]
        [InlineData("HC013-Test-08.docx", 0, false)]
        [InlineData("HC014-RTL-Table-01.docx", 0, true)]
        [InlineData("HC015-Vertical-Spacing-atLeast.docx", 0, false)]
        [InlineData("HC016-Horizontal-Spacing-firstLine.docx", 0, false)]
        [InlineData("HC017-Vertical-Alignment-Cell-01.docx", 0, false)]
        [InlineData("HC018-Vertical-Alignment-Para-01.docx", 0, false)]
        [InlineData("HC019-Hidden-Run.docx", 0, false)]
        [InlineData("HC020-Small-Caps.docx", 0, false)]
        [InlineData("HC021-Symbols.docx", 0, false)]
        [InlineData("HC022-Table-Of-Contents.docx", 0, false)]
        [InlineData("HC023-Hyperlink.docx", 0, false)]
        [InlineData("HC024-Tabs-01.docx", 0, false)]
        [InlineData("HC025-Tabs-02.docx", 0, false)]
        [InlineData("HC026-Tabs-03.docx", 0, false)]
        [InlineData("HC027-Tabs-04.docx", 0, false)]
        [InlineData("HC028-No-Break-Hyphen.docx", 0, false)]
        [InlineData("HC029-Table-Merged-Cells.docx", 0, false)]
        [InlineData("HC030-Content-Controls.docx", 0, false)]
        [InlineData("HC031-Complicated-Document.docx", 0, true)]
        [InlineData("HC032-Named-Color.docx", 0, false)]
        [InlineData("HC033-Run-With-Border.docx", 0, false)]
        [InlineData("HC034-Run-With-Position.docx", 0, false)]
        [InlineData("HC035-Strike-Through.docx", 0, false)]
        [InlineData("HC036-Super-Script.docx", 0, false)]
        [InlineData("HC037-Sub-Script.docx", 0, false)]
        [InlineData("HC038-Conflicting-Border-Weight.docx", 0, false)]
        [InlineData("HC039-Bold.docx", 0, false)]
        [InlineData("HC040-Hyperlink-Fieldcode-01.docx", 0, false)]
        [InlineData("HC041-Hyperlink-Fieldcode-02.docx", 0, false)]
        [InlineData("HC042-Image-Png.docx", 0, false)]
        [InlineData("HC043-Chart.docx", 0, false)]
        [InlineData("HC044-Embedded-Workbook.docx", 0, false)]
        [InlineData("HC045-Italic.docx", 0, false)]
        [InlineData("HC046-BoldAndItalic.docx", 0, false)]
        [InlineData("HC047-No-Section.docx", 0, false)]
        [InlineData("HC048-Excerpt.docx", 0, true)]
        [InlineData("HC049-Borders.docx", 0, false)]
        [InlineData("HC050-Shaded-Text-01.docx", 0, false)]
        [InlineData("HC051-Shaded-Text-02.docx", 0, false)]
        [InlineData("HC052-SmartArt.docx", 0, false)]
        [InlineData("HC053-Headings.docx", 0, false)]
        [InlineData("HC055-GoogleDocsExport.docx", 0, false)]
        [InlineData("HC060-Image-with-Hyperlink.docx", 0, false)]
        [InlineData("HC061-Hyperlink-in-Field.docx", 0, false)]
        [InlineData("Tabs.docx", 0, false)]
        public async Task HC001(string name, int expectedPixelNoise, bool imageSizeMayDiffer)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var settings = new WmlToHtmlConverterSettings(sourceDocx.FullName);

            var oxPtConvertedDestHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-3-OxPt.html")));
            await ConvertToHtml(sourceDocx, oxPtConvertedDestHtml, settings, expectedPixelNoise, imageSizeMayDiffer);
        }

        [Theory]
        [InlineData("HC006-Test-01.docx", 0)]
        public async Task HC002_NoCssClasses(string name, int expectedPixelNoise)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var settings = new WmlToHtmlConverterSettings(sourceDocx.FullName, new ImageHandler(), new TextDummyHandler(), new SymbolHandler(), new BreakHandler(), new FontHandler(), false);

            var oxPtConvertedDestHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-5-OxPt-No-CSS-Classes.html")));
            await ConvertToHtml(sourceDocx, oxPtConvertedDestHtml, settings, expectedPixelNoise, true);
        }

        [Theory]
        [InlineData("HC023-Hyperlink.docx", "href=\"http://example.com/#anchor\"")]
        [InlineData("HC003-Hebrew-02.docx", "<ul ", "<ol ")]
        [InlineData("HC009-Test-04.docx", "<ul ", "<ol ")]
        [InlineData("HC010-Test-05.docx", "<ul ")]
        [InlineData("HC031-Complicated-Document.docx", "<ul ", "<ol ")]
        [InlineData("HC006-Test-01.docx", "<ol ")]
        [InlineData("HC012-Test-07.docx", "<ol ")]
        [InlineData("HC048-Excerpt.docx", "<ol ")]
        [InlineData("HC061-Hyperlink-in-Field.docx", "<ol ")]
        [InlineData("HC053-Headings.docx", "<h1 ", "<h2 ", "<h3 ", "<h4 ", "<h5 ", "<h6 ", "role=\"heading\"", "aria-level=\"7\"", "aria-level=\"8\"", "aria-level=\"9\"")]
        public async Task HC003_ContainsSubstring(string name, params string[] expectedSubstrings)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var settings = new WmlToHtmlConverterSettings(sourceDocx.FullName, new ImageHandler(), new TextDummyHandler(), new SymbolHandler(), new BreakHandler(), new FontHandler(), false);

            var byteArray = await File.ReadAllBytesAsync(sourceDocx.FullName);
            using var memoryStream = new MemoryStream();
            memoryStream.Write(byteArray, 0, byteArray.Length);
            using var wDoc = WordprocessingDocument.Open(memoryStream, true);

            var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
            var htmlString = html.ToString(SaveOptions.DisableFormatting);

            foreach (string expectedSubstring in expectedSubstrings)
            {
                Assert.Contains(expectedSubstring, htmlString);
            }
        }

        private static async Task ConvertToHtml(FileInfo sourceDocx, FileInfo destFileName, WmlToHtmlConverterSettings settings, int expectedPixeNoise, bool imageSizeMayDiffer)
        {
            var byteArray = File.ReadAllBytes(sourceDocx.FullName);
            var expectedRenderdResult = Path.Combine(sourceDocx.Directory.FullName, sourceDocx.Name + "Expectation.png");
            using var memoryStream = new MemoryStream();
            memoryStream.Write(byteArray, 0, byteArray.Length);
            using var wDoc = WordprocessingDocument.Open(memoryStream, true);
            var outputDirectory = destFileName.Directory;
            destFileName = new FileInfo(Path.Combine(outputDirectory.FullName, destFileName.Name));

            var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

            // Note: the XHTML returned by ConvertToHtmlTransform contains objects of type XEntity. PtOpenXmlUtil.cs define the XEntity class. See http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx for detailed explanation. If you further transform the XML tree returned by ConvertToHtmlTransform, you must do it correctly, or entities will not be serialized properly.

            var htmlString = html.ToString(SaveOptions.DisableFormatting);
            File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);

            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                return;
            }

            await AssertRenderedHtmlIsEqual(destFileName.FullName, expectedRenderdResult, expectedPixeNoise, imageSizeMayDiffer);
        }

        internal static async Task AssertRenderedHtmlIsEqual(string actualFilePath, string expectReferenceFilePath, int allowedPixelErrorCount, bool imageSizeMayDiffer)
        {
            var actualFullPath = Path.GetFullPath(actualFilePath);

            Assert.True(File.Exists(actualFullPath), $"actualFilePath not found {actualFullPath}");
            var launchOptions = new LaunchOptions
            {
                Args = new[] { "--force-device-scale-factor=1" },
            };
            await using var chromiumRenderer = await Renderer.CreateAsync(new BrowserFetcher(), launchOptions);
            var pathRasterizedHtml = actualFilePath + ".png";
            await chromiumRenderer.ConvertHtmlToPng(actualFilePath, pathRasterizedHtml);

            await AssertImageIsEqual(pathRasterizedHtml, expectReferenceFilePath, allowedPixelErrorCount, imageSizeMayDiffer);
        }

        internal static async Task AssertImageIsEqual(string actualImagePath, string expectImageFilePath, int allowedPixelErrorCount, bool imageSizeMayDiffer)
        {
            var actualFullPath = Path.GetFullPath(actualImagePath);
            var expectFullPath = Path.GetFullPath(expectImageFilePath);

            Assert.True(File.Exists(actualFullPath), $"actualImagePath not found {actualFullPath}");

            //Uncomment following line to create or update expectation for new test cases
            //File.Copy(actualFullPath, expectFullPath, true);

            Assert.True(File.Exists(expectFullPath), $"ExpectReferenceImagePath not found \n{expectFullPath}\n copy over \n{actualFullPath}\n if this is a new test case.");

            var resizeOption = imageSizeMayDiffer ? ResizeOption.Resize : ResizeOption.DontResize;
            if (Compare.ImagesAreEqual(actualFullPath, expectFullPath, resizeOption, transparencyOptions: TransparencyOptions.CompareAlphaChannel))
            {
                return;
            }

            var allowedDiffInfo = CalcAllowedDiffInfo(actualFullPath, expectFullPath, allowedPixelErrorCount, imageSizeMayDiffer);

            try
            {
                using (var maskImage = Compare.CalcDiffMaskImage(actualFullPath, expectFullPath, ResizeOption.Resize, transparencyOptions: TransparencyOptions.CompareAlphaChannel))
                {
                    var png = maskImage.Encode(SKEncodedImageFormat.Png, 100);
                    await File.WriteAllBytesAsync(allowedDiffInfo.NewDiffImageFileName, png.ToArray());
                }

                // Uncomment following two lines to add an allowed diff file
                //var expectedDirectoryPath = Path.GetDirectoryName(expectFullPath);
                //File.Copy(allowedDiffInfo.NewDiffImageFileName, Path.Combine(expectedDirectoryPath, allowedDiffInfo.NewDiffImageFileName), true);

                if (!imageSizeMayDiffer && !Compare.ImagesHaveEqualSize(actualFullPath, expectFullPath))
                {
                    // Uncomment following line to create or update a allowed diff file
                    //File.Copy(actualFullPath, expectFullPath, true);

                    SaveToGithubActionsPickupTestresultsDirectory(actualFullPath, expectFullPath, allowedDiffInfo.NewDiffImageFileName);
                    Assert.Fail($"Actual dimension differs from expected \nExpected {expectFullPath}\ndiffers to actual {actualFullPath} \nReplace {expectFullPath} with the new value, or configure the test to expect different dimensions.");
                }

                if (allowedDiffInfo.DiffFileExists)
                {
                    var resultWithAllowedDiffPixelErrorCount = int.MaxValue;
                    foreach (var allowedDiffImage in allowedDiffInfo.ExistingDiffImageFilename)
                    {
                        var resultWithAllowedDiff = Compare.CalcDiff(actualFullPath, expectFullPath, allowedDiffImage, resizeOption, transparencyOptions: TransparencyOptions.CompareAlphaChannel);

                        var pixelErrorCountAboveExpectedWithDiff = resultWithAllowedDiff.PixelErrorCount > allowedPixelErrorCount;

                        if (!pixelErrorCountAboveExpectedWithDiff)
                        {
                            return;
                        }
                        resultWithAllowedDiffPixelErrorCount = Math.Min(resultWithAllowedDiffPixelErrorCount, resultWithAllowedDiff.PixelErrorCount);
                    }

                    SaveToGithubActionsPickupTestresultsDirectory(actualFullPath, expectFullPath, allowedDiffInfo.NewDiffImageFileName);
                    Assert.Fail($"Expected PixelErrorCount beyond {allowedPixelErrorCount} but was {resultWithAllowedDiffPixelErrorCount}\nExpected {expectFullPath}\ndiffers to actual {actualFullPath}\n diff is {allowedDiffInfo.NewDiffImageFileName}\n");
                }

                var result = Compare.CalcDiff(actualFullPath, expectFullPath, resizeOption, transparencyOptions: TransparencyOptions.CompareAlphaChannel);

                var pixelErrorCountAboveExpected = result.PixelErrorCount > allowedPixelErrorCount;
                if (pixelErrorCountAboveExpected)
                {
                    SaveToGithubActionsPickupTestresultsDirectory(actualFullPath, expectFullPath, allowedDiffInfo.NewDiffImageFileName);

                    Assert.Fail($"Expected PixelErrorCount less or equal {allowedPixelErrorCount} but was {result.PixelErrorCount}\nExpected {expectFullPath}\ndiffers to actual {actualFullPath}\n Diff is {allowedDiffInfo.NewDiffImageFileName} \nReplace {actualFullPath} with the new value or store the diff as {allowedDiffInfo.ExistingDiffImageFilename.First()}.");
                }
            }
            catch (System.Exception ex) when (!(ex is Xunit.Sdk.FailException))
            {
                SaveToGithubActionsPickupTestresultsDirectory(actualFullPath, expectFullPath, allowedDiffInfo.NewDiffImageFileName);
            }
        }

        private static AllowedDiffInfo CalcAllowedDiffInfo(string actualFullPath, string expectFullPath, int allowedPixelErrorCount, bool imageSizeMayDiffer)
        {
            var osSpecificDiffFileSuffix = RuntimeInformation.IsOSPlatform(OSPlatform.Linux) ? "linux" : "win";
            var newDiffImageFileName = $"{Path.GetFileName(expectFullPath)}.diff.{osSpecificDiffFileSuffix}{DateTime.Now:yyyyMMddHHmmss}.png";

            var directoryPath = Path.GetDirectoryName(expectFullPath);
            var fileNamePattern = $"{Path.GetFileName(expectFullPath)}.diff.{osSpecificDiffFileSuffix}*.png";
            var matchingFiles = Directory.GetFiles(directoryPath, fileNamePattern);

            return new AllowedDiffInfo(matchingFiles.Length > 0, newDiffImageFileName, matchingFiles);
        }

        private static void SaveToGithubActionsPickupTestresultsDirectory(string actualFullPath, string expectFullPath, string newDiffImageFileName)
        {
            var fileName = Path.GetFileName(actualFullPath);
            var expectFullDirectory = Path.GetDirectoryName(expectFullPath);
            var expectFullDirectoryFullPath = Path.GetFullPath(expectFullDirectory);

            var testResultDirectoryActual = Path.Combine(expectFullDirectoryFullPath, "../TestResult/Actual");
            var testResultDirectoryExpected = Path.Combine(expectFullDirectoryFullPath, "../TestResult/Expected");
            var testResultDirectoryDiff = Path.Combine(expectFullDirectoryFullPath, "../TestResult/Diff");
            CreateDirectory(Path.Combine(expectFullDirectoryFullPath, "../TestResult"));
            CreateDirectory(testResultDirectoryActual);
            CreateDirectory(testResultDirectoryExpected);
            CreateDirectory(testResultDirectoryDiff);

            File.Copy(actualFullPath, Path.Combine(testResultDirectoryActual, fileName), true);
            File.Copy(expectFullPath, Path.Combine(testResultDirectoryExpected, fileName), true);
            var newDiffImageFullPath = Path.GetFullPath(newDiffImageFileName);
            var destFileName = Path.Combine(testResultDirectoryDiff, Path.GetFileName(newDiffImageFullPath));
            File.Copy(newDiffImageFullPath, destFileName, true);

            static void CreateDirectory(string testResultDirectory)
            {
                if (!Directory.Exists(testResultDirectory))
                {
                    Directory.CreateDirectory(testResultDirectory);
                }
            }
        }
    }
}