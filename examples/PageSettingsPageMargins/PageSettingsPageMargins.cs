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

            // by default there's already a worksheet named "Sheet1"
            sl.AddWorksheet("Sheet2");
            sl.AddWorksheet("Sheet3");

            SLPageSettings ps = new SLPageSettings();

            ps.SetWideMargins();
            sl.SetPageSettings(ps, "Sheet1");

            // the same SLPageSettings variable can be reused
            // because only the page margins setting is changed.
            // Otherwise, use separate variables.

            ps.SetNarrowMargins();
            sl.SetPageSettings(ps, "Sheet2");

            ps.TopMargin = 5.43;
            ps.BottomMargin = 3.45;
            // any unassigned margin takes on the "Normal" margin default
            // However, in this case, the last margin setting was "Narrow",
            // so the unassigned margins takes on the "Narrow" margin settings.

            // without a given sheet name, the currently selected sheet is used
            sl.SetPageSettings(ps);

            sl.SaveAs("PageSettingsPageMargins.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
