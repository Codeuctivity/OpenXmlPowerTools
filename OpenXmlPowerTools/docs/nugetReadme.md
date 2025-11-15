# OpenXmlPowerTools

[![.NET build and test](https://github.com/Codeuctivity/OpenXmlPowerTools/actions/workflows/dotnet.yml/badge.svg)](https://github.com/Codeuctivity/OpenXmlPowerTools/actions/workflows/dotnet.yml) [![Nuget](https://img.shields.io/nuget/v/Codeuctivity.OpenXmlPowerTools.svg)](https://www.nuget.org/packages/Codeuctivity.OpenXmlPowerTools/)

## Focus of this fork

- Linux, Windows and MacOs support
- Conversion of DOCX to HTML/CSS.

## Example - Convert DOCX to HTML

``` csharp
var sourceDocxFileContent = File.ReadAllBytes("./source.docx");
using var memoryStream = new MemoryStream();
await memoryStream.WriteAsync(sourceDocxFileContent, 0, sourceDocxFileContent.Length);
using var wordProcessingDocument = WordprocessingDocument.Open(memoryStream, true);
var settings = new WmlToHtmlConverterSettings("htmlPageTitle");
var html = WmlToHtmlConverter.ConvertToHtml(wordProcessingDocument, settings);
var htmlString = html.ToString(SaveOptions.DisableFormatting);
File.WriteAllText("./target.html", htmlString, Encoding.UTF8);
```