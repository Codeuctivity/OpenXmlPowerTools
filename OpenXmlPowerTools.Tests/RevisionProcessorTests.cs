﻿using Codeuctivity.OpenXmlPowerTools;
using System;
using System.IO;
using Xunit;

namespace Codeuctivity.Tests
{
    public class RpTests
    {
        [Theory]
        [InlineData("RP/RP002-Deleted-Text.docx")]
        [InlineData("RP/RP003-Inserted-Text.docx")]
        [InlineData("RP/RP004-Deleted-Text-in-CC.docx")]
        [InlineData("RP/RP005-Deleted-Paragraph-Mark.docx")]
        [InlineData("RP/RP006-Inserted-Paragraph-Mark.docx")]
        [InlineData("RP/RP007-Multiple-Deleted-Para-Mark.docx")]
        [InlineData("RP/RP008-Multiple-Inserted-Para-Mark.docx")]
        [InlineData("RP/RP009-Deleted-Table-Row.docx")]
        [InlineData("RP/RP010-Inserted-Table-Row.docx")]
        [InlineData("RP/RP011-Multiple-Deleted-Rows.docx")]
        [InlineData("RP/RP012-Multiple-Inserted-Rows.docx")]
        [InlineData("RP/RP013-Deleted-Math-Control-Char.docx")]
        [InlineData("RP/RP014-Inserted-Math-Control-Char.docx")]
        [InlineData("RP/RP015-MoveFrom-MoveTo.docx")]
        [InlineData("RP/RP016-Deleted-CC.docx")]
        [InlineData("RP/RP017-Inserted-CC.docx")]
        [InlineData("RP/RP018-MoveFrom-MoveTo-CC.docx")]
        [InlineData("RP/RP019-Deleted-Field-Code.docx")]
        [InlineData("RP/RP020-Inserted-Field-Code.docx")]
        [InlineData("RP/RP021-Inserted-Numbering-Properties.docx")]
        [InlineData("RP/RP022-NumberingChange.docx")]
        [InlineData("RP/RP023-NumberingChange.docx")]
        [InlineData("RP/RP024-ParagraphMark-rPr-Change.docx")]
        [InlineData("RP/RP025-Paragraph-Props-Change.docx")]
        [InlineData("RP/RP026-NumberingChange.docx")]
        [InlineData("RP/RP027-Change-Section.docx")]
        [InlineData("RP/RP028-Table-Grid-Change.docx")]
        [InlineData("RP/RP029-Table-Row-Props-Change.docx")]
        [InlineData("RP/RP030-Table-Row-Props-Change.docx")]
        [InlineData("RP/RP031-Table-Prop-Change.docx")]
        [InlineData("RP/RP032-Table-Prop-Change.docx")]
        [InlineData("RP/RP033-Table-Prop-Ex-Change.docx")]
        [InlineData("RP/RP034-Deleted-Cells.docx")]
        [InlineData("RP/RP035-Inserted-Cells.docx")]
        [InlineData("RP/RP036-Vert-Merged-Cells.docx")]
        [InlineData("RP/RP037-Changed-Style-Para-Props.docx")]
        [InlineData("RP/RP038-Inserted-Paras-at-End.docx")]
        [InlineData("RP/RP039-Inserted-Paras-at-End.docx")]
        [InlineData("RP/RP040-Deleted-Paras-at-End.docx")]
        [InlineData("RP/RP041-Cell-With-Empty-Paras-at-End.docx")]
        [InlineData("RP/RP042-Deleted-Para-Mark-at-End.docx")]
        [InlineData("RP/RP043-MERGEFORMAT-Field-Code.docx")]
        [InlineData("RP/RP044-MERGEFORMAT-Field-Code.docx")]
        [InlineData("RP/RP045-One-and-Half-Deleted-Lines-at-End.docx")]
        [InlineData("RP/RP046-Consecutive-Deleted-Ranges.docx")]
        [InlineData("RP/RP047-Inserted-and-Deleted-Paragraph-Mark.docx")]
        [InlineData("RP/RP048-Deleted-Inserted-Para-Mark.docx")]
        [InlineData("RP/RP049-Deleted-Para-Before-Table.docx")]
        [InlineData("RP/RP050-Deleted-Footnote.docx")]
        [InlineData("RP/RP052-Deleted-Para-Mark.docx")]
        public void RP001(string name)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceFi = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var baselineAcceptedFi = new FileInfo(Path.Combine(sourceDir.FullName, name.Replace(".docx", "-Accepted.docx")));
            var baselineRejectedFi = new FileInfo(Path.Combine(sourceDir.FullName, name.Replace(".docx", "-Rejected.docx")));

            var sourceWml = new WmlDocument(sourceFi.FullName);
            var afterRejectingWml = RevisionProcessor.RejectRevisions(sourceWml);
            var afterAcceptingWml = RevisionProcessor.AcceptRevisions(sourceWml);

            var processedAcceptedFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceFi.Name.Replace(".docx", "-Accepted.docx")));
            afterAcceptingWml.SaveAs(processedAcceptedFi.FullName);

            var processedRejectedFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceFi.Name.Replace(".docx", "-Rejected.docx")));
            afterRejectingWml.SaveAs(processedRejectedFi.FullName);

            // create batch file to copy properly processed documents to the TestFiles directory.
            while (true)
            {
                try
                {
                    var batchFileName = "Copy-Gen-Files-To-TestFiles.bat";
                    var batchFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, batchFileName));
                    var batch = "";
                    batch += "copy " + processedAcceptedFi.FullName + " " + baselineAcceptedFi.FullName + Environment.NewLine;
                    batch += "copy " + processedRejectedFi.FullName + " " + baselineRejectedFi.FullName + Environment.NewLine;
                    if (batchFi.Exists)
                    {
                        File.AppendAllText(batchFi.FullName, batch);
                    }
                    else
                    {
                        File.WriteAllText(batchFi.FullName, batch);
                    }
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }
    }
}