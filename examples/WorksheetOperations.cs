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
            // SpreadsheetLight works on the idea of a currently selected worksheet.
            // So there's always a worksheet selected, just like you always only
            // have one worksheet active when in Excel.
            SLDocument sl = new SLDocument();
            // At this point, there's already 1 worksheet available for use.

            // If you're opening an existing spreadsheet, the first available
            // worksheet in the spreadsheet is selected. And that worksheet may not
            // be what you want. So for efficiency (and possibly less hair-tearing),
            // you can use this to make sure you got the correct worksheet.
            // SLDocument sl = new SLDocument("YourFile.xlsx", "WorksheetToSelect");

            // The first worksheet of a new spreadsheet is named "Sheet1",
            // but use the constant to future-proof your application.
            sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "HulkSmash");

            for (int i = 1; i < 20; ++i)
            {
                for (int j = 1; j < 15; ++j)
                {
                    sl.SetCellValue(i, j, string.Format("R{0}C{1}", i, j));
                }
            }

            sl.AddWorksheet("Superman");
            sl.SetCellValue(2, 2, "I'm Superman");

            sl.CopyWorksheet("HulkSmash", "Crater");

            sl.AddWorksheet("Blackhole");
            sl.SetCellValue(3, 3, "This is gonna be deleted...");

            sl.SelectWorksheet("Crater");
            sl.SetCellValue(25, 2, "That's a big crater");

            sl.DeleteWorksheet("Blackhole");

            // Superman is currently in 2nd place, after HulkSmash.
            // This moves Superman to position 1.
            sl.MoveWorksheet("Superman", 1);

            sl.SaveAs("WorksheetOperations.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
