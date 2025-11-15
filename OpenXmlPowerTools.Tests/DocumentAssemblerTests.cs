using Codeuctivity.OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Xunit;
using System.Globalization;

namespace Codeuctivity.Tests
{
    public class DaTests
    {
        [Theory]
        [InlineData("DA001-TemplateDocument.docx", "DA-Data.xml", false)]
        [InlineData("DA002-TemplateDocument.docx", "DA-DataNotHighValueCust.xml", false)]
        [InlineData("DA003-Select-XPathFindsNoData.docx", "DA-Data.xml", true)]
        [InlineData("DA004-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml", false)]
        [InlineData("DA005-SelectRowData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA006-SelectTestValue-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA007-SelectRepeatingData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA008-TableElementWithNoTable.docx", "DA-Data.xml", true)]
        [InlineData("DA009-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA010-InvalidXml.docx", "DA-Data.xml", true)]
        [InlineData("DA011-SchemaError.docx", "DA-Data.xml", true)]
        [InlineData("DA012-OtherMarkupTypes.docx", "DA-Data.xml", true)]
        [InlineData("DA013-Runs.docx", "DA-Data.xml", false)]
        [InlineData("DA014-TwoRuns-NoValuesSelected.docx", "DA-Data.xml", true)]
        [InlineData("DA015-TwoRunsXmlExceptionInFirst.docx", "DA-Data.xml", true)]
        [InlineData("DA016-TwoRunsSchemaErrorInSecond.docx", "DA-Data.xml", true)]
        [InlineData("DA017-FiveRuns.docx", "DA-Data.xml", true)]
        [InlineData("DA018-SmartQuotes.docx", "DA-Data.xml", false)]
        [InlineData("DA019-RunIsEntireParagraph.docx", "DA-Data.xml", false)]
        [InlineData("DA020-TwoRunsAndNoOtherContent.docx", "DA-Data.xml", true)]
        [InlineData("DA021-NestedRepeat.docx", "DA-DataNestedRepeat.xml", false)]
        [InlineData("DA022-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA023-RepeatWOEndRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA026-InvalidRootXmlElement.docx", "DA-Data.xml", true)]
        [InlineData("DA027-XPathErrorInPara.docx", "DA-Data.xml", true)]
        [InlineData("DA028-NoPrototypeRow.docx", "DA-Data.xml", true)]
        [InlineData("DA029-NoDataForCell.docx", "DA-Data.xml", true)]
        [InlineData("DA030-TooMuchDataForCell.docx", "DA-TooMuchDataForCell.xml", true)]
        [InlineData("DA031-CellDataInAttributes.docx", "DA-CellDataInAttributes.xml", true)]
        [InlineData("DA032-TooMuchDataForConditional.docx", "DA-TooMuchDataForConditional.xml", true)]
        [InlineData("DA033-ConditionalOnAttribute.docx", "DA-ConditionalOnAttribute.xml", false)]
        [InlineData("DA034-HeaderFooter.docx", "DA-Data.xml", false)]
        [InlineData("DA035-SchemaErrorInRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA036-SchemaErrorInConditional.docx", "DA-Data.xml", true)]
        [InlineData("DA100-TemplateDocument.docx", "DA-Data.xml", false)]
        [InlineData("DA101-TemplateDocument.docx", "DA-Data.xml", true)]
        [InlineData("DA102-TemplateDocument.docx", "DA-Data.xml", true)]
        [InlineData("DA201-TemplateDocument.docx", "DA-Data.xml", false)]
        [InlineData("DA202-TemplateDocument.docx", "DA-DataNotHighValueCust.xml", false)]
        [InlineData("DA203-Select-XPathFindsNoData.docx", "DA-Data.xml", true)]
        [InlineData("DA204-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml", false)]
        [InlineData("DA205-SelectRowData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA206-SelectTestValue-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA207-SelectRepeatingData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA209-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA210-InvalidXml.docx", "DA-Data.xml", true)]
        [InlineData("DA211-SchemaError.docx", "DA-Data.xml", true)]
        [InlineData("DA212-OtherMarkupTypes.docx", "DA-Data.xml", true)]
        [InlineData("DA213-Runs.docx", "DA-Data.xml", false)]
        [InlineData("DA214-TwoRuns-NoValuesSelected.docx", "DA-Data.xml", true)]
        [InlineData("DA215-TwoRunsXmlExceptionInFirst.docx", "DA-Data.xml", true)]
        [InlineData("DA216-TwoRunsSchemaErrorInSecond.docx", "DA-Data.xml", true)]
        [InlineData("DA217-FiveRuns.docx", "DA-Data.xml", true)]
        [InlineData("DA218-SmartQuotes.docx", "DA-Data.xml", false)]
        [InlineData("DA219-RunIsEntireParagraph.docx", "DA-Data.xml", false)]
        [InlineData("DA220-TwoRunsAndNoOtherContent.docx", "DA-Data.xml", true)]
        [InlineData("DA221-NestedRepeat.docx", "DA-DataNestedRepeat.xml", false)]
        [InlineData("DA222-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA223-RepeatWOEndRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA226-InvalidRootXmlElement.docx", "DA-Data.xml", true)]
        [InlineData("DA227-XPathErrorInPara.docx", "DA-Data.xml", true)]
        [InlineData("DA228-NoPrototypeRow.docx", "DA-Data.xml", true)]
        [InlineData("DA229-NoDataForCell.docx", "DA-Data.xml", true)]
        [InlineData("DA230-TooMuchDataForCell.docx", "DA-TooMuchDataForCell.xml", true)]
        [InlineData("DA231-CellDataInAttributes.docx", "DA-CellDataInAttributes.xml", true)]
        [InlineData("DA232-TooMuchDataForConditional.docx", "DA-TooMuchDataForConditional.xml", true)]
        [InlineData("DA233-ConditionalOnAttribute.docx", "DA-ConditionalOnAttribute.xml", false)]
        [InlineData("DA234-HeaderFooter.docx", "DA-Data.xml", false)]
        [InlineData("DA235-Crashes.docx", "DA-Content-List.xml", false)]
        [InlineData("DA236-Page-Num-in-Footer.docx", "DA-Content-List.xml", false)]
        [InlineData("DA237-SchemaErrorInRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA238-SchemaErrorInConditional.docx", "DA-Data.xml", true)]
        [InlineData("DA239-RunLevelCC-Repeat.docx", "DA-Data.xml", false)]
        [InlineData("DA250-ConditionalWithRichXPath.docx", "DA250-Address.xml", false)]
        [InlineData("DA251-EnhancedTables.docx", "DA-Data.xml", false)]
        [InlineData("DA252-Table-With-Sum.docx", "DA-Data.xml", false)]
        [InlineData("DA253-Table-With-Sum-Run-Level-CC.docx", "DA-Data.xml", false)]
        [InlineData("DA254-Table-With-XPath-Sum.docx", "DA-Data.xml", false)]
        [InlineData("DA255-Table-With-XPath-Sum-Run-Level-CC.docx", "DA-Data.xml", false)]
        [InlineData("DA256-NoInvalidDocOnErrorInRun.docx", "DA-Data.xml", true)]
        [InlineData("DA257-OptionalRepeat.docx", "DA-Data.xml", false)]
        [InlineData("DA258-ContentAcceptsCharsAsXPathResult.docx", "DA-Data.xml", false)]
        [InlineData("DA259-MultiLineContents.docx", "DA-Data.xml", false)]
        [InlineData("DA260-RunLevelRepeat.docx", "DA-Data.xml", false)]
        [InlineData("DA261-RunLevelConditional.docx", "DA-Data.xml", false)]
        [InlineData("DA262-ConditionalNotMatch.docx", "DA-Data.xml", false)]
        [InlineData("DA263-ConditionalNotMatch.docx", "DA-DataSmallCustomer.xml", false)]
        [InlineData("DA264-InvalidRunLevelRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA265-RunLevelRepeatWithWhiteSpaceBefore.docx", "DA-Data.xml", false)]
        [InlineData("DA266-RunLevelRepeat-NoData.docx", "DA-Data.xml", true)]
        public void DA101(string name, string data, bool err)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, data));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmldata = XElement.Load(dataFile.FullName);

            var afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out var returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            afterAssembling.SaveAs(assembledDocx.FullName);

            using (var ms = new MemoryStream())
            {
                ms.Write(afterAssembling.DocumentByteArray, 0, afterAssembling.DocumentByteArray.Length);
                using var wDoc = WordprocessingDocument.Open(ms, true);
                var v = new OpenXmlValidator();
                var valErrors = v.Validate(wDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));

                var sb = new StringBuilder();
                foreach (var item in valErrors.Select(r => r.Description).OrderBy(t => t).Distinct())
                {
                    sb.Append(item).Append(Environment.NewLine);
                }
                var z = sb.ToString();
                Console.WriteLine(z);

                Assert.Empty(valErrors);
            }

            Assert.Equal(err, returnedTemplateError);
        }

        [Theory]
        [InlineData("DA259-MultiLineContents.docx", "DA-Data.xml", false)]
        public void DA259(string name, string data, bool err)
        {
            DA101(name, data, err);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            var afterAssembling = new WmlDocument(assembledDocx.FullName);
            var brCount = afterAssembling.MainDocumentPart
                            .Element(W.body)
                            .Elements(W.p).ElementAt(1)
                            .Elements(W.r)
                            .Elements(W.br).Count();
            Assert.Equal(4, brCount);
        }

        [Theory]
        [InlineData("DA024-TrackedRevisions.docx", "DA-Data.xml")]
        public void DA102_Throws(string name, string data)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, data));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmldata = XElement.Load(dataFile.FullName);

            WmlDocument afterAssembling;
            Assert.Throws<OpenXmlPowerToolsException>(() =>
                {
                    afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out var returnedTemplateError);
                });
        }

        [Theory]
        [InlineData("DA025-TemplateDocument.docx", "DA-Data.xml", false)]
        public void DA103_UseXmlDocument(string name, string data, bool err)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, data));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmldata = new XmlDocument();
            xmldata.Load(dataFile.FullName);

            var afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out var returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            afterAssembling.SaveAs(assembledDocx.FullName);

            using (var ms = new MemoryStream())
            {
                ms.Write(afterAssembling.DocumentByteArray, 0, afterAssembling.DocumentByteArray.Length);
                using var wDoc = WordprocessingDocument.Open(ms, true);
                var v = new OpenXmlValidator();
                var valErrors = v.Validate(wDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));
                Assert.Empty(valErrors);
            }

            Assert.Equal(err, returnedTemplateError);
        }

        [Fact]
        public void AssembleDocument_ImageMetadataCreatesImageParts()
        {
            var template = CreateTemplateDocument("DA-ImageTemplate.docx",
                "<# <Image Select=\"Image[1]\" /> #>",
                "<# <Image Select=\"Image[2]\" /> #>");
            var data = new XElement("Images",
                new XElement("Image", TinyPngBase64),
                new XElement("Image", TinyPngBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.False(templateError);

            using var ms = new MemoryStream(assembled.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var mainPart = wDoc.MainDocumentPart;

            Assert.Equal(2, mainPart.ImageParts.Count());

            var docPrIds = mainPart
                .GetXDocument()
                .Descendants(WP.docPr)
                .Select(d => (int)d.Attribute("id"))
                .ToList();
            Assert.Equal(new[] { 1, 2 }, docPrIds);

            var blipIds = mainPart
                .GetXDocument()
                .Descendants(A.blip)
                .Select(b => (string)b.Attribute(R.embed))
                .ToList();
            Assert.Equal(2, blipIds.Count);
            Assert.All(blipIds, id => Assert.False(string.IsNullOrEmpty(id)));
        }

        [Fact]
        public void AssembleDocument_InvalidBase64RaisesTemplateError()
        {
            var template = CreateTemplateDocument("DA-ImageInvalidTemplate.docx", "<# <Image Select=\"Image[1]\" /> #>");
            var data = new XElement("Images", new XElement("Image", "not-base64"));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.True(templateError);

            using var ms = new MemoryStream(assembled.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var text = string.Concat(wDoc.MainDocumentPart.GetXDocument().Descendants(W.t).Select(t => (string)t));
            Assert.Contains("Image:", text);
        }

        [Fact]
        public void AssembleDocument_OptionalImageIsSkippedWhenMissing()
        {
            var template = CreateTemplateDocument("DA-OptionalImageTemplate.docx", "<# <Image Select=\"Missing\" Optional=\"true\" /> #>");
            var data = new XElement("Images");

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.False(templateError);

            using var ms = new MemoryStream(assembled.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var drawings = wDoc.MainDocumentPart.GetXDocument().Descendants(W.drawing);
            Assert.Empty(drawings);
            Assert.Empty(wDoc.MainDocumentPart.ImageParts);
        }

        [Fact]
        public void AssembleDocument_ImageWidthScalesHeight()
        {
            var template = CreateTemplateDocument("DA-ImageWidth.docx", "<# <Image Select=\"Image[1]\" Width=\"1in\" /> #>");
            var data = new XElement("Images", new XElement("Image", LargeSamplePngBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.False(templateError);

            var (paragraph, extent) = ExtractImageParagraph(assembled);
            Assert.Null(paragraph.Element(W.pPr));
            Assert.Equal(EmusPerInch.ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cx"));
            Assert.Equal((EmusPerInch / 2).ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cy"));
        }

        [Fact]
        public void AssembleDocument_ImageHeightScalesWidth()
        {
            var template = CreateTemplateDocument("DA-ImageHeight.docx", "<# <Image Select=\"Image[1]\" Height=\"1in\" /> #>");
            var data = new XElement("Images", new XElement("Image", LargeSamplePngBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.False(templateError);

            var (paragraph, extent) = ExtractImageParagraph(assembled);
            Assert.Null(paragraph.Element(W.pPr));
            Assert.Equal((EmusPerInch * 2).ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cx"));
            Assert.Equal(EmusPerInch.ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cy"));
        }

        [Fact]
        public void AssembleDocument_ImageMaxWidthCentersParagraph()
        {
            var template = CreateTemplateDocument("DA-ImageMaxWidth.docx", "<# <Image Select=\"Image[1]\" Align=\"center\" MaxWidth=\"3in\" /> #>");
            var data = new XElement("Images", new XElement("Image", LargeSamplePngBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.False(templateError);

            var (paragraph, extent) = ExtractImageParagraph(assembled);
            Assert.Equal("center", (string)paragraph.Element(W.pPr)?.Element(W.jc)?.Attribute(W.val));
            var expectedWidth = 3 * EmusPerInch;
            var expectedHeight = expectedWidth / 2;
            Assert.Equal(expectedWidth.ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cx"));
            Assert.Equal(expectedHeight.ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cy"));
        }

        [Fact]
        public void AssembleDocument_ImageMaxHeightJustifiesParagraph()
        {
            var template = CreateTemplateDocument("DA-ImageMaxHeight.docx", "<# <Image Select=\"Image[1]\" Align=\"justify\" MaxHeight=\"150px\" /> #>");
            var data = new XElement("Images", new XElement("Image", TallPngBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.False(templateError);

            var (paragraph, extent) = ExtractImageParagraph(assembled);
            Assert.Equal("both", (string)paragraph.Element(W.pPr)?.Element(W.jc)?.Attribute(W.val));
            var maxHeight = 150 * EmusPerPixel;
            var expectedWidth = (200 * EmusPerPixel) * (maxHeight / (400 * EmusPerPixel));
            Assert.Equal(expectedWidth.ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cx"));
            Assert.Equal(maxHeight.ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cy"));
        }

        [Fact]
        public void AssembleDocument_ImageMaxHeightAndWidthAppliesLargestConstraint()
        {
            var template = CreateTemplateDocument("DA-ImageMaxBoth.docx", "<# <Image Select=\"Image[1]\" Align=\"right\" Width=\"4in\" MaxHeight=\"1in\" /> #>");
            var data = new XElement("Images", new XElement("Image", LargeSamplePngBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.False(templateError);

            var (paragraph, extent) = ExtractImageParagraph(assembled);
            Assert.Equal("right", (string)paragraph.Element(W.pPr)?.Element(W.jc)?.Attribute(W.val));
            Assert.Equal((2 * EmusPerInch).ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cx"));
            Assert.Equal(EmusPerInch.ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cy"));
        }

        [Fact]
        public void AssembleDocument_ImageInvalidAlignProducesTemplateError()
        {
            var template = CreateTemplateDocument("DA-ImageInvalidAlign.docx", "<# <Image Select=\"Image[1]\" Align=\"diagonal\" /> #>");
            var data = new XElement("Images", new XElement("Image", LargeSamplePngBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.True(templateError);

            var text = GetDocumentText(assembled);
            Assert.Contains("Align attribute must be one of Left, Center, Right, or Justify.", text);
        }

        [Fact]
        public void AssembleDocument_ImageInvalidWidthReportsError()
        {
            var template = CreateTemplateDocument("DA-ImageInvalidWidth.docx", "<# <Image Select=\"Image[1]\" Width=\"abc\" /> #>");
            var data = new XElement("Images", new XElement("Image", LargeSamplePngBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.True(templateError);
            var text = GetDocumentText(assembled);
            Assert.Contains("Unable to parse length 'abc'.", text);
        }

        [Fact]
        public void AssembleDocument_ImageZeroHeightReportsError()
        {
            var template = CreateTemplateDocument("DA-ImageZeroHeight.docx", "<# <Image Select=\"Image[1]\" Height=\"0in\" /> #>");
            var data = new XElement("Images", new XElement("Image", LargeSamplePngBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.True(templateError);
            var text = GetDocumentText(assembled);
            Assert.Contains("Length value '0in' must be greater than zero.", text);
        }

        [Fact]
        public void AssembleDocument_ImageGifDimensionsFallback()
        {
            var template = CreateTemplateDocument("DA-ImageGifFallback.docx", "<# <Image Select=\"Image[1]\" /> #>");
            var data = new XElement("Images", new XElement("Image", TruncatedGifBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.False(templateError);

            var (paragraph, extent) = ExtractImageParagraph(assembled);
            Assert.Null(paragraph.Element(W.pPr));
            Assert.Equal((200 * EmusPerPixel).ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cx"));
            Assert.Equal((80 * EmusPerPixel).ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cy"));
        }

        [Fact]
        public void AssembleDocument_ImageRespectsMaxDimensionsAndAlign()
        {
            var template = CreateTemplateDocument("DA-ImageConstraintTemplate.docx", "<# <Image Select=\"Image[1]\" Align=\"center\" MaxWidth=\"300px\" /> #>");
            var data = new XElement("ImageData", new XElement("Image", LargeSamplePngBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.False(templateError);

            using var ms = new MemoryStream(assembled.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var paragraph = wDoc.MainDocumentPart.GetXDocument().Descendants(W.p).FirstOrDefault(p => p.Descendants(W.drawing).Any());
            Assert.NotNull(paragraph);
            var jc = paragraph!.Element(W.pPr)?.Element(W.jc)?.Attribute(W.val)?.Value;
            Assert.Equal("center", jc);

            var extent = paragraph.Descendants(WP.extent).First();
            Assert.Equal(2857500d.ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cx"));
            Assert.Equal(1428750d.ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cy"));
        }

        [Fact]
        public void AssembleDocument_ImageUsesExplicitDimensions()
        {
            var template = CreateTemplateDocument("DA-ImageExplicitSizeTemplate.docx", "<# <Image Select=\"Image[1]\" Align=\"right\" Width=\"2in\" Height=\"1in\" /> #>");
            var data = new XElement("ImageData", new XElement("Image", LargeSamplePngBase64));

            var assembled = DocumentAssembler.AssembleDocument(template, data, out var templateError);
            Assert.False(templateError);

            using var ms = new MemoryStream(assembled.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var paragraph = wDoc.MainDocumentPart.GetXDocument().Descendants(W.p).FirstOrDefault(p => p.Descendants(W.drawing).Any());
            Assert.NotNull(paragraph);
            var jc = paragraph!.Element(W.pPr)?.Element(W.jc)?.Attribute(W.val)?.Value;
            Assert.Equal("right", jc);

            var extent = paragraph.Descendants(WP.extent).First();
            Assert.Equal(1828800d.ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cx"));
            Assert.Equal(914400d.ToString("0", CultureInfo.InvariantCulture), (string)extent.Attribute("cy"));
        }

        private static WmlDocument CreateTemplateDocument(string fileName, params string[] paragraphTexts)
        {
            using var ms = new MemoryStream();
            using (var wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = wDoc.AddMainDocumentPart();
                var body = new Body(paragraphTexts.Select(text =>
                    new Paragraph(
                        new Run(
                            new Text(text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }))));
                mainPart.Document = new Document(body);
                mainPart.Document.Save();
            }

            return new WmlDocument(fileName, ms.ToArray());
        }

        private static (XElement Paragraph, XElement Extent) ExtractImageParagraph(WmlDocument document)
        {
            using var ms = new MemoryStream();
            ms.Write(document.DocumentByteArray, 0, document.DocumentByteArray.Length);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var main = wDoc.MainDocumentPart.GetXDocument();
            var paragraph = main.Descendants(W.p).First(p => p.Descendants(W.drawing).Any());
            var extent = paragraph.Descendants(WP.extent).First();
            return (paragraph, extent);
        }

        private static string GetDocumentText(WmlDocument document)
        {
            using var ms = new MemoryStream();
            ms.Write(document.DocumentByteArray, 0, document.DocumentByteArray.Length);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var main = wDoc.MainDocumentPart.GetXDocument();
            return string.Concat(main.Descendants(W.t).Select(t => (string)t));
        }

        private const double EmusPerInch = 914400d;
        private const double EmusPerPixel = EmusPerInch / 96d;
        private const string TinyPngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
        private const string LargeSamplePngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAZAAAADICAIAAABJdyC1AAACvElEQVR4nO3WsQ3DQBAEsXvB/bf8akHZYQyygo0Ge2buABQ82wMAvhIsIEOwgAzBAjIEC8gQLCBDsIAMwQIyBAvIECwgQ7CADMECMgQLyBAsIEOwgAzBAjIEC8gQLCBDsIAMwQIyBAvIECwgQ7CADMECMgQLyBAsIEOwgAzBAjIEC8gQLCBDsIAMwQIyBAvIECwgQ7CADMECMgQLyPhtDyi5c7Yn8J/O3O0JDR4WkCFYQIZgARmCBWQIFpAhWECGYAEZggVkCBaQIVhAhmABGYIFZAgWkCFYQIZgARmCBWQIFpAhWECGYAEZggVkCBaQIVhAhmABGYIFZAgWkCFYQIZgARmCBWQIFpAhWECGYAEZggVkCBaQIVhAhmABGYIFZAgWkCFYQIZgARmCBWQIFpAhWECGYAEZggVkCBaQIVhAhmABGYIFZAgWkCFYQIZgARmCBWQIFpAhWECGYAEZggVkCBaQIVhAhmABGYIFZAgWkCFYQIZgARmCBWQIFpAhWECGYAEZggVkCBaQIVhAhmABGYIFZAgWkCFYQIZgARmCBWScmbu9AeATDwvIECwgQ7CADMECMgQLyBAsIEOwgAzBAjIEC8gQLCBDsIAMwQIyBAvIECwgQ7CADMECMgQLyBAsIEOwgAzBAjIEC8gQLCBDsIAMwQIyBAvIECwgQ7CADMECMgQLyBAsIEOwgAzBAjIEC8gQLCBDsIAMwQIyBAvIECwgQ7CADMECMgQLyBAsIEOwgAzBAjIEC8gQLCBDsIAMwQKm4gVVxAWP2c47WwAAAABJRU5ErkJggg==";

        private const string WidePngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAZAAAADICAIAAABJdyC1AAACuUlEQVR4nO3UMQ7CQBAEwT3EvxEvXz/BZKalqniCifrs7gAUvGfmnO/TNwBu7H5edxuAfyFYQIZgARmCBWQIFpAhWECGYAEZggVkCBaQIVhAhmABGYIFZAgWkCFYQIZgARmCBWQIFpAhWECGYAEZggVkCBaQIVhAhmABGYIFZAgWkCFYQIZgARmCBWQIFpAhWECGYAEZggVkCBaQIVhAhmABGYIFZAgWkCFYQIZgARmCBWQIFpAhWECGYAEZggVkCBaQIVhAhmABGYIFZAgWkCFYQIZgARmCBWQIFpAhWECGYAEZggVkCBaQIVhAhmABGYIFFzMzu2y8A5u4PQZkIj89BEMEAAAAASUVORK5CYII=";
        private const string TallPngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAMgAAAGQCAIAAABkkLjnAAAEF0lEQVR4nO3S0QkCURAEwX1i3mLke0lcI3hVAQzz0Wd3B+72npnzPbfv8mT72devP/CfhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkzu42y8yTXVnDDBu1Y983AAAAAElFTkSuQmCC";
        private const string TruncatedGifBase64 = "R0lGODlhyABQAA==";
        private static readonly List<string> s_ExpectedErrors = new List<string>()
        {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddVBand' attribute is not declared.",
        };
    }
}
