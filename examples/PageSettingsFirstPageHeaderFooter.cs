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
            SLDocument sl = new SLDocument();

            SLPageSettings ps = new SLPageSettings();
            // It's a good practice to get the current page settings
            // and then work on it. But there's nothing in the worksheet
            // anyway, so it doesn't really matter.
            // SLPageSettings ps = sl.GetPageSettings();

            ps.OddHeaderText = "An odd header";
            ps.OddFooterText = "An odd footer";

            ps.FirstHeaderText = "First page header";
            ps.FirstFooterText = "First page footer";

            ps.DifferentFirstPage = true;

            sl.SetPageSettings(ps);

            sl.SaveAs("PageSettingsFirstPageHeaderFooter.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
