using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            SLDocument sl = new SLDocument();

            sl.SetCellValue(1, 1, "This is frozen");
            // this freezes the top 4 rows and the 5 left-most columns
            sl.FreezePanes(4, 5);

            sl.AddWorksheet("Sheet2");
            sl.SetCellValue(1, 1, "This is split");
            // this splits the worksheet.
            // The top-left pane is 6 rows high and 7 columns wide.
            // "true" to show row and column headers.
            sl.SplitPanes(6, 7, true);

            sl.SaveAs("FreezeSplitPanes.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
