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

            sl.SetCellValue("A1", "I am fond of fonts.");

            SLStyle style = sl.CreateStyle();
            style.Font.FontName = "Harrington";
            style.Font.FontSize = 18;
            style.Font.FontColor = System.Drawing.Color.Blue;
            style.Font.Bold = true;
            style.Font.Italic = true;
            style.Font.Strike = true;
            style.Font.Underline = UnderlineValues.Double;
            sl.SetCellStyle("A1", style);

            sl.SetCellValue(3, 3, "The font of all fonts should be a thing.");

            // certain font properties can be directly accessed with shortcut functions
            style = sl.CreateStyle();
            // this uses the minor font of the theme at 24 points
            style.SetFont(FontSchemeValues.Minor, 24);
            // this uses the accent 4 colour, and brightens it by 30%
            style.SetFontColor(SLThemeColorIndexValues.Accent4Color, 0.3);
            style.SetFontBold(true);
            style.SetFontUnderline(UnderlineValues.Single);
            sl.SetCellStyle(3, 3, style);

            sl.SaveAs("Font.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
