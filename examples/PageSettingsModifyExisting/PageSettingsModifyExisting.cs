using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            SLDocument sl = new SLDocument("PageSettingsExisting.xlsx", "Sheet3");

            // Sheet1 has a landscape orientation, scaled 150% and a green tab color
            SLPageSettings ps = sl.GetPageSettings("Sheet1");

            ps.SetWideMargins();

            // Sheet2 only has "Normal" page margins
            sl.SetPageSettings(ps, "Sheet2");
            // But now will have Sheet1 page settings, and a "Wide" margin.

            // Just to show the operation only happens between Sheet1 and Sheet2,
            // here's setting a cell value
            sl.SetCellValue(3, 3, "Sheet3 not modified with page settings!");

            sl.SaveAs("PageSettingsModifyExisting.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
