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

            SLPageSettings ps = sl.GetPageSettings();
            ps.View = SheetViewValues.PageLayout;
            // SpreadsheetLight by default has a worksheet ready for you named "Sheet1".
            // This sets to the currently selected worksheet, which is "Sheet1".
            sl.SetPageSettings(ps);

            sl.SetCellValue(2, 2, "This worksheet is in page layout mode.");

            sl.SaveAs("PageLayout.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
