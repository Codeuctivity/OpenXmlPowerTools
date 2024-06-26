﻿using Codeuctivity.OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

internal class ListItemRetriever01
{
    private class XmlStackItem
    {
        public XElement Element;
        public int[] LevelNumbers;
    }

    private static void Main()
    {
        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        using var wDoc =
            WordprocessingDocument.Open("../../NumberedListTest.docx", false);
        var abstractNumId = 0;
        var xml = ConvertDocToXml(wDoc, abstractNumId);
        Console.WriteLine(xml);
        xml.Save(Path.Combine(tempDi.FullName, "Out.xml"));
    }

    private static XElement ConvertDocToXml(WordprocessingDocument wDoc, int abstractNumId)
    {
        var xd = wDoc.MainDocumentPart.GetXDocument();

        // First, call RetrieveListItem so that all paragraphs are initialized with ListItemInfo
        var firstParagraph = xd.Descendants(W.p).FirstOrDefault();
        var listItem = ListItemRetriever.RetrieveListItem(wDoc, firstParagraph);

        var xml = new XElement("Root");
        var current = new Stack<XmlStackItem>();
        current.Push(
            new XmlStackItem()
            {
                Element = xml,
                LevelNumbers = Array.Empty<int>(),
            });
        foreach (var paragraph in xd.Descendants(W.p))
        {
            // The following does not take into account documents that have tracked revisions.
            // As necessary, call RevisionAccepter.AcceptRevisions before converting to XML.
            var text = paragraph.Descendants(W.t).Select(t => (string)t).StringConcatenate();
            var lii =
                paragraph.Annotation<ListItemRetriever.ListItemInfo>();
            if (lii.IsListItem && lii.AbstractNumId == abstractNumId)
            {
                var levelNums =
                    paragraph.Annotation<ListItemRetriever.LevelNumbers>();
                if (levelNums.LevelNumbersArray.Length == current.Peek().LevelNumbers.Length)
                {
                    current.Pop();
                    var levelNumsForThisIndent = levelNums.LevelNumbersArray;
                    var levelText = levelNums
                        .LevelNumbersArray
                        .Select(l => l.ToString() + ".")
                        .StringConcatenate()
                        .TrimEnd('.');
                    var newCurrentElement = new XElement("Indent",
                        new XAttribute("Level", levelText));
                    current.Peek().Element.Add(newCurrentElement);
                    current.Push(
                        new XmlStackItem()
                        {
                            Element = newCurrentElement,
                            LevelNumbers = levelNumsForThisIndent,
                        });
                    current.Peek().Element.Add(new XElement("Heading", text));
                }
                else if (levelNums.LevelNumbersArray.Length > current.Peek().LevelNumbers.Length)
                {
                    for (var i = current.Peek().LevelNumbers.Length;
                        i < levelNums.LevelNumbersArray.Length;
                        i++)
                    {
                        var levelNumsForThisIndent = levelNums
                            .LevelNumbersArray
                            .Take(i + 1)
                            .ToArray();
                        var levelText = levelNums
                            .LevelNumbersArray
                            .Select(l => l.ToString() + ".")
                            .StringConcatenate()
                            .TrimEnd('.');
                        var newCurrentElement = new XElement("Indent",
                            new XAttribute("Level", levelText));
                        current.Peek().Element.Add(newCurrentElement);
                        current.Push(
                            new XmlStackItem()
                            {
                                Element = newCurrentElement,
                                LevelNumbers = levelNumsForThisIndent,
                            });
                        current.Peek().Element.Add(new XElement("Heading", text));
                    }
                }
                else if (levelNums.LevelNumbersArray.Length < current.Peek().LevelNumbers.Length)
                {
                    for (var i = current.Peek().LevelNumbers.Length;
                        i > levelNums.LevelNumbersArray.Length;
                        i--)
                    {
                        current.Pop();
                    }

                    current.Pop();
                    var levelNumsForThisIndent = levelNums.LevelNumbersArray;
                    var levelText = levelNums
                        .LevelNumbersArray
                        .Select(l => l.ToString() + ".")
                        .StringConcatenate()
                        .TrimEnd('.');
                    var newCurrentElement = new XElement("Indent",
                        new XAttribute("Level", levelText));
                    current.Peek().Element.Add(newCurrentElement);
                    current.Push(
                        new XmlStackItem()
                        {
                            Element = newCurrentElement,
                            LevelNumbers = levelNumsForThisIndent,
                        });
                    current.Peek().Element.Add(new XElement("Heading", text));
                }
            }
            else
            {
                current.Peek().Element.Add(new XElement("Paragraph", text));
            }
        }
        return xml;
    }
}