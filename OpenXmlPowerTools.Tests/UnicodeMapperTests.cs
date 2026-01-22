using Codeuctivity.OpenXmlPowerTools;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace Codeuctivity.Tests
{
    public class UnicodeMapperTests
    {
        [Fact]
        public void CanStringifyRunAndTextElements()
        {
            const string textValue = "Hello World!";
            var textElement = new XElement(W.t, textValue);
            var runElement = new XElement(W.r, textElement);
            var formattedRunElement = new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), textElement);

            Assert.Equal(textValue, UnicodeMapper.RunToString(textElement));
            Assert.Equal(textValue, UnicodeMapper.RunToString(runElement));
            Assert.Equal(textValue, UnicodeMapper.RunToString(formattedRunElement));
        }

        [Fact]
        public void CanStringifySpecialElements()
        {
            Assert.Equal(UnicodeMapper.CarriageReturn,
                UnicodeMapper.RunToString(new XElement(W.cr)).First());
            Assert.Equal(UnicodeMapper.CarriageReturn,
                UnicodeMapper.RunToString(new XElement(W.br)).First());
            Assert.Equal(UnicodeMapper.FormFeed,
                UnicodeMapper.RunToString(new XElement(W.br, new XAttribute(W.type, "page"))).First());
            Assert.Equal(UnicodeMapper.NonBreakingHyphen,
                UnicodeMapper.RunToString(new XElement(W.noBreakHyphen)).First());
            Assert.Equal(UnicodeMapper.SoftHyphen,
                UnicodeMapper.RunToString(new XElement(W.softHyphen)).First());
            Assert.Equal(UnicodeMapper.HorizontalTabulation,
                UnicodeMapper.RunToString(new XElement(W.tab)).First());
        }

        [Fact]
        public void CanCreateRunChildElementsFromSpecialCharacters()
        {
            Assert.Equal(W.br, UnicodeMapper.CharToRunChild(UnicodeMapper.CarriageReturn).Name);
            Assert.Equal(W.noBreakHyphen, UnicodeMapper.CharToRunChild(UnicodeMapper.NonBreakingHyphen).Name);
            Assert.Equal(W.softHyphen, UnicodeMapper.CharToRunChild(UnicodeMapper.SoftHyphen).Name);
            Assert.Equal(W.tab, UnicodeMapper.CharToRunChild(UnicodeMapper.HorizontalTabulation).Name);

            var element = UnicodeMapper.CharToRunChild(UnicodeMapper.FormFeed);
            Assert.Equal(W.br, element.Name);
            Assert.Equal("page", element.Attribute(W.type).Value);

            Assert.Equal(W.br, UnicodeMapper.CharToRunChild('\r').Name);
        }

        [Fact]
        public void CanCreateCoalescedRuns()
        {
            const string textString = "This is only text.";
            const string mixedString = "First\tSecond\tThird";

            var textRuns = UnicodeMapper.StringToCoalescedRunList(textString, null);
            var mixedRuns = UnicodeMapper.StringToCoalescedRunList(mixedString, null);

            Assert.Single(textRuns);
            Assert.Equal(5, mixedRuns.Count);

            Assert.Equal("First", mixedRuns.Elements(W.t).Skip(0).First().Value);
            Assert.Equal("Second", mixedRuns.Elements(W.t).Skip(1).First().Value);
            Assert.Equal("Third", mixedRuns.Elements(W.t).Skip(2).First().Value);
        }

        [Fact]
        public void CanMapSymbols()
        {
            var sym1 = new XElement(W.sym,
                new XAttribute(W.font, "Wingdings"),
                new XAttribute(W._char, "F028"));
            var charFromSym1 = UnicodeMapper.SymToChar(sym1);
            var symFromChar1 = UnicodeMapper.CharToRunChild(charFromSym1);

            var sym2 = new XElement(W.sym,
                new XAttribute(W._char, "F028"),
                new XAttribute(W.font, "Wingdings"));
            var charFromSym2 = UnicodeMapper.SymToChar(sym2);

            var sym3 = new XElement(W.sym,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(W.font, "Wingdings"),
                new XAttribute(W._char, "F028"));
            var charFromSym3 = UnicodeMapper.SymToChar(sym3);

            var sym4 = new XElement(W.sym,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(W.font, "Webdings"),
                new XAttribute(W._char, "F028"));
            var charFromSym4 = UnicodeMapper.SymToChar(sym4);
            var symFromChar4 = UnicodeMapper.CharToRunChild(charFromSym4);

            Assert.Equal(charFromSym1, charFromSym2);
            Assert.Equal(charFromSym1, charFromSym3);
            Assert.NotEqual(charFromSym1, charFromSym4);

            Assert.Equal("F028", symFromChar1.Attribute(W._char).Value);
            Assert.Equal("Wingdings", symFromChar1.Attribute(W.font).Value);

            Assert.Equal("F028", symFromChar4.Attribute(W._char).Value);
            Assert.Equal("Webdings", symFromChar4.Attribute(W.font).Value);
        }

        [Fact]
        public void CanStringifySymbols()
        {
            var charFromSym1 = UnicodeMapper.SymToChar("Wingdings", '\uF028');
            var charFromSym2 = UnicodeMapper.SymToChar("Wingdings", 0xF028);
            var charFromSym3 = UnicodeMapper.SymToChar("Wingdings", "F028");

            var symFromChar1 = UnicodeMapper.CharToRunChild(charFromSym1);
            var symFromChar2 = UnicodeMapper.CharToRunChild(charFromSym2);
            var symFromChar3 = UnicodeMapper.CharToRunChild(charFromSym3);

            Assert.Equal(charFromSym1, charFromSym2);
            Assert.Equal(charFromSym1, charFromSym3);

            Assert.Equal(symFromChar1.ToString(SaveOptions.None), symFromChar2.ToString(SaveOptions.None));
            Assert.Equal(symFromChar1.ToString(SaveOptions.None), symFromChar3.ToString(SaveOptions.None));
        }

        private const string LastRenderedPageBreakXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t>ThisIsAParagraphContainingNoNaturalLi</w:t>
      </w:r>
      <w:r>
        <w:lastRenderedPageBreak/>
        <w:t>neBreaksSoTheLineBreakIsForced.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";

        [Fact]
        public void IgnoresTemporaryLayoutMarkers()
        {
            XDocument partDocument = XDocument.Parse(LastRenderedPageBreakXmlString);
            XElement p = partDocument.Descendants(W.p).Last();
            string actual = p.Descendants(W.r)
                .Select(UnicodeMapper.RunToString)
                .StringConcatenate();
            // p.Value is "the concatenated text content of this element", which
            // (in THIS test case, which does not feature any symbols or special
            // characters) should exactly match the output of UnicodeMapper:
            Assert.Equal(p.Value, actual);
        }

        private const string PreserveSpacingXmlString =
@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space=""preserve"">The following space is retained: </w:t>
      </w:r>
      <w:r>
        <w:t>but this one is not: </w:t>
      </w:r>
      <w:r>
        <w:t xml:space=""preserve"">. Similarly these two lines should have only a space between them: </w:t>
      </w:r>
      <w:r>
        <w:t>
          Line 1!
Line 2!
        </w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";

        [Fact]
        public void HonorsXmlSpace()
        {
            // This somewhat rudimentary test is superceded by TreatsXmlSpaceLikeWord() below,
            // but it has been left in to provide a simple/direct illustration of a couple of
            // the specific test cases covered by that more extensive suite.
            XDocument partDocument = XDocument.Parse(PreserveSpacingXmlString);
            XElement p = partDocument.Descendants(W.p).Last();
            string innerText = p.Descendants(W.r)
                .Select(UnicodeMapper.RunToString)
                .StringConcatenate();
            Assert.Equal(@"The following space is retained: but this one is not:. Similarly these two lines should have only a space between them: Line 1! Line 2!", innerText);
        }

        // Verifies that UnicodeMapper.RunToString interprets whitespace in <w:t> elements
        // exactly the way Microsoft Word does, including honoring xml:space="preserve".
        // This is essential because RunToString is used by higher‑level features
        // (OpenXmlRegex, DocumentAssembler, etc.) that rely on its output to reflect the
        // text an end‑user would actually see and edit in Word.
        //
        // Word accepts a wide range of “valid” DOCX input, but it normalizes that input
        // into a canonical form when displaying or saving the document. These tests
        // compare RunToString’s output against Word’s canonicalized output to ensure
        // that whitespace is treated as semantic content in the same way Word treats it.
        [Fact]
        public void TreatsXmlSpaceLikeWord()
        {
            var sourceDir = new System.IO.DirectoryInfo("../../../../TestFiles/");
            // Test document: crafted to include many whitespace patterns that Word accepts as valid input
            var testDoc = new System.IO.FileInfo(System.IO.Path.Combine(sourceDir.FullName, "UM-Whitespace-test.docx"));
            var testWmlDoc = new WmlDocument(testDoc.FullName);
            var testParagraphs = testWmlDoc.MainDocumentPart
                            .Element(W.body)
                            .Elements(W.p).ToList();
            // Canonical document: the same test document after being opened and saved by Word,
            // representing Word's own normalized interpretation of that whitespace
            var expectedDoc = new System.IO.FileInfo(System.IO.Path.Combine(sourceDir.FullName, "UM-Whitespace-Word-saved.docx"));
            var expectedWmlDoc = new WmlDocument(expectedDoc.FullName);
            var expectedParagraphs = expectedWmlDoc.MainDocumentPart
                            .Element(W.body)
                            .Elements(W.p).ToList();
            // Iterate through pairs of paragraphs (test name, test content, expected result)
            for (int i = 0; i < testParagraphs.Count - 1; i += 2)
            {
                var testNameParagraph = testParagraphs[i];
                var testContentParagraph = testParagraphs[i + 1];
                // Get the test name from the first paragraph
                var testName = testNameParagraph.Descendants(W.t)
                    .Select(t => (string)t)
                    .StringConcatenate();
                // Get the actual result by calling UnicodeMapper.RunToString on the test content runs
                var actualResult = testContentParagraph.Descendants(W.r)
                    .Select(UnicodeMapper.RunToString)
                    .StringConcatenate();
                // Find corresponding expected result paragraph (same index in expected document)
                var expectedResult = ExtractExpectedFromWord(expectedParagraphs[i + 1]);
                Assert.True(
                    expectedResult == actualResult,
                    $"Test '{testName}' failed. Expected: [{expectedResult}] Actual: [{actualResult}]"
                );
            }
        }

        // Extracts the expected text from Word’s canonicalized output for the whitespace tests.
        // This helper intentionally handles *only* the constructs that Word emits in the saved
        // version of UM-whitespace-test.docx:
        //   • <w:t>      → literal text
        //   • <w:tab/>   → '\t'
        //   • <w:lastRenderedPageBreak/> (intentionally ignored)
        // If any other run-level element appears, it means Word has emitted something this test
        // was not designed to handle, and the test fails loudly. This prevents the helper
        // from drifting toward reimplementing UnicodeMapper.RunToString.
        private static string ExtractExpectedFromWord(XElement p)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var run in p.Elements(W.r))
            {
                foreach (var child in run.Elements())
                {
                    if (child.Name == W.t)
                    {
                        sb.Append((string)child);
                    }
                    else if (child.Name == W.tab)
                    {
                        sb.Append('\t');
                    }
                    else if (child.Name != W.lastRenderedPageBreak)
                    {
                        throw new System.InvalidOperationException(
                            $"Unexpected element <{child.Name.LocalName}> encountered in expected Word output.");
                    }
                }
            }
            return sb.ToString();
        }
    }
}