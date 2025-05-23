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

            // you can set text straight in
            ps.OddHeaderText = "This is a header";

            SLFont ft = sl.CreateFont();
            ft.SetFont("Impact", 16);
            ps.AppendOddFooter(ft, "This is page ");

            ps.AppendOddFooter(SLHeaderFooterFormatCodeValues.PageNumber);

            // to undo the font settings from before
            ps.AppendOddFooter(SLHeaderFooterFormatCodeValues.ResetFont);

            ps.AppendOddFooter(" of ");

            ft = sl.CreateFont();
            ft.SetFont(FontSchemeValues.Major, 18);
            ft.SetFontThemeColor(SLThemeColorIndexValues.Accent1Color);
            // we use an empty string here so the font style is set.
            // Format codes follow the previous style, unless a style reset command
            // is given
            ps.AppendOddFooter(ft, "");

            ps.AppendOddFooter(SLHeaderFooterFormatCodeValues.NumberOfPages);

            sl.SetPageSettings(ps);

            sl.SaveAs("PageSettingsHeaderFooter.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
