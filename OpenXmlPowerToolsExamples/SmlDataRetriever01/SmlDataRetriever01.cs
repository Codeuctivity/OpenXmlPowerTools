﻿using Codeuctivity.OpenXmlPowerTools;
using System;
using System.IO;

namespace SmlDataRetriever01
{
    internal class SmlDataRetriever01
    {
        private static void Main()
        {
            var fi = new FileInfo("../../SampleSpreadsheet.xlsx");

            // Retrieve range from Sheet1
            var data = SmlDataRetriever.RetrieveRange(fi.FullName, "Sheet1", "A1:C3");
            Console.WriteLine(data);

            // Retrieve entire sheet
            data = SmlDataRetriever.RetrieveSheet(fi.FullName, "Sheet1");
            Console.WriteLine(data);

            // Retrieve table
            data = SmlDataRetriever.RetrieveTable(fi.FullName, "VehicleTable");
            Console.WriteLine(data);
        }
    }
}