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

            sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "RedSheet");
            sl.AddWorksheet("GreenSheet");
            sl.AddWorksheet("BlueSheet");

            SLPageSettings ps = new SLPageSettings();
            // It's a good practice to get the current page settings
            // and then work on it. But there's nothing in the worksheet
            // anyway, so it doesn't really matter.
            // SLPageSettings ps = sl.GetPageSettings();

            ps.TabColor = System.Drawing.Color.Red;
            sl.SetPageSettings(ps, "RedSheet");

            // the same SLPageSettings variable can be reused
            // because only the tab color setting is changed.
            // Otherwise, use separate variables.

            ps.TabColor = System.Drawing.Color.Green;
            sl.SetPageSettings(ps, "GreenSheet");

            // the default office theme has accent 1 as a bluish colour
            ps.SetTabColor(SLThemeColorIndexValues.Accent1Color);
            // without a given sheet name, the currently selected sheet is used
            sl.SetPageSettings(ps);

            sl.SaveAs("PageSettingsTabColor.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
