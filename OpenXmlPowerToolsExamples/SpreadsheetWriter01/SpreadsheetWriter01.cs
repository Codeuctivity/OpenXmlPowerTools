using Codeuctivity.OpenXmlPowerTools;
using System;
using System.IO;

namespace SpreadsheetWriter01
{
    internal class Program
    {
        private static void Main()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            var wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new() {
                        Name = "MyFirstSheet",
                        TableName = "NamesAndRates",
                        ColumnHeadings = new CellDfn[]
                        {
                            new() {
                                Value = "Name",
                                Bold = true,
                            },
                            new() {
                                Value = "Age",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Left,
                            },
                            new() {
                                Value = "Rate",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Left,
                            }
                        },
                        Rows = new RowDfn[]
                        {
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "Eric",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = 50,
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)45.00,
                                        FormatCode = "0.00",
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "Bob",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = 42,
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)78.00,
                                        FormatCode = "0.00",
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(tempDi.FullName, "Test1.xlsx"), wb);
        }
    }
}