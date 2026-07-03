using Codeuctivity.OpenXmlPowerTools;
using System;
using System.IO;

namespace SpreadsheetWriter02
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
                        ColumnHeadings = new CellDfn[]
                        {
                            new() {
                                Value = "DataType",
                                Bold = true,
                            },
                            new() {
                                Value = "Value",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Right,
                            },
                        },
                        Rows = new RowDfn[]
                        {
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Boolean,
                                        Value = true,
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Boolean,
                                        Value = false,
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "String",
                                    },
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "A String",
                                        HorizontalCellAlignment = HorizontalCellAlignment.Right,
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "int",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = 100,
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "int?",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = (int?)100,
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "int? (is null)",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = null,
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "uint",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = (uint)101,
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "long",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = long.MaxValue,
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "float",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = (float)123.45,
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "double",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = 123.45,
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.String,
                                        Value = "decimal",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)123.45,
                                    },
                                }
                            },
                            new() {
                                Cells = new CellDfn[]
                                {
                                    new() {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8),
                                        FormatCode = "mm-dd-yy",
                                    },
                                    new() {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 9),
                                        FormatCode = "mm-dd-yy",
                                        Bold = true,
                                        HorizontalCellAlignment = HorizontalCellAlignment.Center,
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(tempDi.FullName, "Test2.xlsx"), wb);
        }
    }
}