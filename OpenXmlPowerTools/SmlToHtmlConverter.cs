﻿using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Codeuctivity.OpenXmlPowerTools
{
    public partial class SmlDocument
    {
        public XElement ConvertToHtml(SmlToHtmlConverterSettings htmlConverterSettings, string tableName)
        {
            return SmlToHtmlConverter.ConvertTableToHtml(this, htmlConverterSettings, tableName);
        }

        public XElement ConvertTableToHtml(string tableName)
        {
            var settings = new SmlToHtmlConverterSettings();
            return SmlToHtmlConverter.ConvertTableToHtml(this, settings, tableName);
        }
    }

    public class SmlToHtmlConverterSettings
    {
        public string PageTitle { get; set; }
        public string CssClassPrefix { get; set; }
        public bool FabricateCssClasses { get; set; }
        public string GeneralCss { get; set; }
        public string AdditionalCss { get; set; }

        public SmlToHtmlConverterSettings()
        {
            PageTitle = "";
            CssClassPrefix = "pt-";
            FabricateCssClasses = true;
            GeneralCss = "span { white-space: pre-wrap; }";
            AdditionalCss = "";
        }

        public SmlToHtmlConverterSettings(SmlToHtmlConverterSettings htmlConverterSettings)
        {
            PageTitle = htmlConverterSettings.PageTitle;
            CssClassPrefix = htmlConverterSettings.CssClassPrefix;
            FabricateCssClasses = htmlConverterSettings.FabricateCssClasses;
            GeneralCss = htmlConverterSettings.GeneralCss;
            AdditionalCss = htmlConverterSettings.AdditionalCss;
        }
    }

    public static class SmlToHtmlConverter
    {
        public static XElement ConvertTableToHtml(SmlDocument smlDoc, SmlToHtmlConverterSettings settings, string tableName)
        {
            using var ms = new MemoryStream();
            ms.Write(smlDoc.DocumentByteArray, 0, smlDoc.DocumentByteArray.Length);
            using var sDoc = SpreadsheetDocument.Open(ms, false);
            var rangeXml = SmlDataRetriever.RetrieveTable(sDoc, tableName);
            var xhtml = ConvertToHtmlInternal(sDoc, settings, rangeXml);
            return xhtml;
        }

        public static XElement ConvertTableToHtml(SpreadsheetDocument sDoc, SmlToHtmlConverterSettings settings, string tableName)
        {
            var rangeXml = SmlDataRetriever.RetrieveTable(sDoc, tableName);
            var xhtml = ConvertToHtmlInternal(sDoc, settings, rangeXml);
            return xhtml;
        }

        private static XElement ConvertToHtmlInternal(SpreadsheetDocument sDoc, SmlToHtmlConverterSettings htmlConverterSettings, XElement rangeXml)
        {
            var xhtml = (XElement)ConvertToHtmlTransform(sDoc, htmlConverterSettings, rangeXml);

            ReifyStylesAndClasses(htmlConverterSettings, xhtml);

            // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
            // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
            // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
            // for detailed explanation.
            //
            // If you further transform the XML tree returned by ConvertToHtmlTransform, you
            // must do it correctly, or entities will not be serialized properly.

            return xhtml;
        }

        private static XNode ConvertToHtmlTransform(SpreadsheetDocument sDoc, SmlToHtmlConverterSettings htmlConverterSettings, XNode node)
        {
            if (node is XElement element)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => ConvertToHtmlTransform(sDoc, htmlConverterSettings, n)));
            }
            return node;
        }

        private static void ReifyStylesAndClasses(SmlToHtmlConverterSettings htmlConverterSettings, XElement xhtml)
        {
            if (htmlConverterSettings.FabricateCssClasses)
            {
                var usedCssClassNames = new HashSet<string>();
                var elementsThatNeedClasses = xhtml
                    .DescendantsAndSelf()
                    .Select(d => new
                    {
                        Element = d,
                        Styles = d.Annotation<Dictionary<string, string>>(),
                    })
                    .Where(z => z.Styles != null);
                var augmented = elementsThatNeedClasses
                    .Select(p => new
                    {
                        p.Element,
                        p.Styles,
                        StylesString = p.Element.Name.LocalName + "|" + p.Styles.OrderBy(k => k.Key).Select(s => string.Format("{0}: {1};", s.Key, s.Value)).StringConcatenate(),
                    })
                    .GroupBy(p => p.StylesString)
                    .ToList();
                var classCounter = 1000000;
                var sb = new StringBuilder();
                sb.Append(Environment.NewLine);
                foreach (var grp in augmented)
                {
                    string classNameToUse;
                    var firstOne = grp.First();
                    var styles = firstOne.Styles;
                    if (styles.ContainsKey("PtStyleName"))
                    {
                        classNameToUse = htmlConverterSettings.CssClassPrefix + styles["PtStyleName"];
                        if (usedCssClassNames.Contains(classNameToUse))
                        {
                            classNameToUse = htmlConverterSettings.CssClassPrefix +
                                styles["PtStyleName"] + "-" +
                                classCounter.ToString().Substring(1);
                            classCounter++;
                        }
                    }
                    else
                    {
                        classNameToUse = htmlConverterSettings.CssClassPrefix +
                            classCounter.ToString().Substring(1);
                        classCounter++;
                    }
                    usedCssClassNames.Add(classNameToUse);
                    sb.Append(firstOne.Element.Name.LocalName + "." + classNameToUse + " {" + Environment.NewLine);
                    foreach (var st in firstOne.Styles.Where(s => s.Key != "PtStyleName"))
                    {
                        var s = "    " + st.Key + ": " + st.Value + ";" + Environment.NewLine;
                        sb.Append(s);
                    }
                    sb.Append("}" + Environment.NewLine);
                    var classAtt = new XAttribute("class", classNameToUse);
                    foreach (var gc in grp)
                    {
                        gc.Element.Add(classAtt);
                    }
                }
                var styleValue = htmlConverterSettings.GeneralCss + sb + htmlConverterSettings.AdditionalCss;

                SetStyleElementValue(xhtml, styleValue);
            }
            else
            {
                // Previously, the h:style element was not added at this point. However,
                // at least the General CSS will contain important settings.
                SetStyleElementValue(xhtml, htmlConverterSettings.GeneralCss + htmlConverterSettings.AdditionalCss);

                foreach (var d in xhtml.DescendantsAndSelf())
                {
                    var style = d.Annotation<Dictionary<string, string>>();
                    if (style == null)
                    {
                        continue;
                    }

                    var styleValue =
                        style
                        .Where(p => p.Key != "PtStyleName")
                        .OrderBy(p => p.Key)
                        .Select(e => string.Format("{0}: {1};", e.Key, e.Value))
                        .StringConcatenate();
                    var st = new XAttribute("style", styleValue);
                    if (d.Attribute("style") != null)
                    {
                        d.Attribute("style").Value += styleValue;
                    }
                    else
                    {
                        d.Add(st);
                    }
                }
            }
        }

        private static void SetStyleElementValue(XElement xhtml, string styleValue)
        {
            var styleElement = xhtml
                .Descendants(Xhtml.style)
                .FirstOrDefault();
            if (styleElement != null)
            {
                styleElement.Value = styleValue;
            }
            else
            {
                styleElement = new XElement(Xhtml.style, styleValue);
                var head = xhtml.Element(Xhtml.head);
                if (head != null)
                {
                    head.Add(styleElement);
                }
            }
        }

        private static readonly Dictionary<string, string> FontFallback = new Dictionary<string, string>()
        {
            { "Arial", @"'{0}', 'sans-serif'" },
            { "Arial Narrow", @"'{0}', 'sans-serif'" },
            { "Arial Rounded MT Bold", @"'{0}', 'sans-serif'" },
            { "Arial Unicode MS", @"'{0}', 'sans-serif'" },
            { "Baskerville Old Face", @"'{0}', 'serif'" },
            { "Berlin Sans FB", @"'{0}', 'sans-serif'" },
            { "Berlin Sans FB Demi", @"'{0}', 'sans-serif'" },
            { "Calibri Light", @"'{0}', 'sans-serif'" },
            { "Gill Sans MT", @"'{0}', 'sans-serif'" },
            { "Gill Sans MT Condensed", @"'{0}', 'sans-serif'" },
            { "Lucida Sans", @"'{0}', 'sans-serif'" },
            { "Lucida Sans Unicode", @"'{0}', 'sans-serif'" },
            { "Segoe UI", @"'{0}', 'sans-serif'" },
            { "Segoe UI Light", @"'{0}', 'sans-serif'" },
            { "Segoe UI Semibold", @"'{0}', 'sans-serif'" },
            { "Tahoma", @"'{0}', 'sans-serif'" },
            { "Trebuchet MS", @"'{0}', 'sans-serif'" },
            { "Verdana", @"'{0}', 'sans-serif'" },
            { "Book Antiqua", @"'{0}', 'serif'" },
            { "Bookman Old Style", @"'{0}', 'serif'" },
            { "Californian FB", @"'{0}', 'serif'" },
            { "Cambria", @"'{0}', 'serif'" },
            { "Constantia", @"'{0}', 'serif'" },
            { "Garamond", @"'{0}', 'serif'" },
            { "Lucida Bright", @"'{0}', 'serif'" },
            { "Lucida Fax", @"'{0}', 'serif'" },
            { "Palatino Linotype", @"'{0}', 'serif'" },
            { "Times New Roman", @"'{0}', 'serif'" },
            { "Wide Latin", @"'{0}', 'serif'" },
            { "Courier New", @"'{0}'" },
            { "Lucida Console", @"'{0}'" },
        };
    }
}