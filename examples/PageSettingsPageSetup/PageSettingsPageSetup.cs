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

            // this should result in displaying "#NAME?" on the worksheet.
            // However, on print preview, it should be "#N/A" (see settings below)
            sl.SetCellValue(2, 2, "=SPREADSHEETLIGHT");

            SLPageSettings ps = new SLPageSettings();

            ps.Orientation = OrientationValues.Landscape;

            // 120% of normal page size
            ps.ScalePage(120);

            ps.PaperSize = SLPaperSizeValues.A4Paper;
            ps.HorizontalDpi = 300;
            ps.VerticalDpi = 300;

            ps.PrintGridLines = true;
            ps.BlackAndWhite = true;
            ps.Draft = true;
            ps.PrintHeadings = true;

            ps.CellComments = CellCommentsValues.AtEnd;
            ps.Errors = PrintErrorValues.NA;

            ps.PageOrder = PageOrderValues.OverThenDown;

            sl.SetPageSettings(ps);

            sl.SaveAs("PageSettingsPageSetup.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
