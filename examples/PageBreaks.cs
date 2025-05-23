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

            // this inserts a page break above row 5,
            // and a page break to the left of column 7
            // Use a negative number to ignore the respective row or column.
            sl.InsertPageBreak(5, 7);

            // this inserts a page break above row 12
            sl.InsertPageBreak(12, -1);

            // this removes the page break above row 5 (if it exists).
            //sl.RemovePageBreak(5, -1);

            // for when you're tired of your things breaking...
            //sl.RemoveAllPageBreaks();

            // this should be in page 2
            sl.SetCellValue(7, 2, "I'm on page 2! Even Steven.");

            // this should be in page 6
            sl.SetCellValue(15, 10, "I'm on page 6! Pick up sticks!");

            // You don't have to set this part.
            // But if you don't, you have to manually change to
            // PageBreakPreview mode in Excel to see the page breaks.
            SLPageSettings ps = sl.GetPageSettings();
            ps.View = SheetViewValues.PageBreakPreview;
            sl.SetPageSettings(ps);

            sl.SaveAs("PageBreaks.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
