# UnicodeMapper whitespace fixtures

`UnicodeMapperTests.TreatsXmlSpaceLikeWord` compares each crafted input paragraph
with the literal text and tabs in the corresponding paragraph saved by Microsoft
Word. Each case has a name paragraph followed by a content paragraph. Each document
ends with one blank paragraph. The tests check paragraph counts and matching names
so a missing or reordered case cannot silently pass.

The original `UM-Whitespace-test.docx` and `UM-Whitespace-Word-saved.docx` contain
107 cases from the January 2026 investigation. They are unchanged.

The supplementary `UM-Whitespace-Unicode-test.docx` and
`UM-Whitespace-Unicode-Word-saved.docx` contain 43 cases. The expected document was
opened and saved without content edits using Microsoft Word for Mac 16.114 on
October 21, 2026. These are synthetic documents; no customer document was used.

The supplementary cases cover:

- Leading, trailing, standalone, and mixed-boundary NBSP with `xml:space` absent,
  `default`, and `preserve`.
- An NBSP separator at either side of a run boundary and in a standalone run.
- All other XML-valid characters recognized as whitespace by .NET: U+0085,
  U+1680, U+2000 through U+200A, U+2028, U+2029, U+202F, U+205F, and U+3000.
  Each is surrounded by ordinary XML whitespace to distinguish trimming the XML
  whitespace from trimming the Unicode character.
- Ordinary spaces, tabs, CR, and CRLF, with preservation absent and present.
  CR and LF use character references so the XML parser does not normalize them
  before Word or UnicodeMapper sees them. Literal XML line endings in the original
  fixture are normalized by XML parsing and do not make that distinction.

Word preserves the tested Unicode characters, trims only boundary space/tab/CR/LF
when preservation is absent or `default`, converts internal line endings to spaces,
and retains tabs when preservation is enabled. These observations cover the listed
inputs and Word version, not every possible whitespace or formatting context.

`UM-Whitespace-Unicode-test.xml` is the exact `word/document.xml` payload in the
supplementary input DOCX, retained for review and reproduction. To regenerate the
input from the repository root using Python's standard library:

```python
from pathlib import Path
from zipfile import ZipFile

root = Path("TestFiles")
xml = (root / "UM-Whitespace-Unicode-test.xml").read_bytes()
with ZipFile(root / "UM-Whitespace-test.docx") as source:
    with ZipFile(root / "UM-Whitespace-Unicode-test.docx", "w") as target:
        for entry in source.infolist():
            data = xml if entry.filename == "word/document.xml" else source.read(entry.filename)
            target.writestr(entry, data)
```

Keep the crafted input unchanged. Open it in Word and use Save As to create the
separate `UM-Whitespace-Unicode-Word-saved.docx`. Do not manufacture expected output
with UnicodeMapper or another word processor. Review any changes in the Word-saved
text before updating expectations.
