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

            SLStyle style1 = sl.CreateStyle();
            style1.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Accent2Color, SLThemeColorIndexValues.Accent4Color);

            SLStyle style2 = sl.CreateStyle();
            style2.SetFont(FontSchemeValues.Minor, 18);
            style2.Fill.SetGradient(SLGradientShadingStyleValues.Corner1, SLThemeColorIndexValues.Accent1Color, SLThemeColorIndexValues.Accent6Color);

            // set row 4 with 1st style
            sl.SetRowStyle(4, style1);
            // set rows 5 through 8 with 2nd style
            sl.SetRowStyle(5, 8, style2);

            SLStyle style3 = sl.CreateStyle();
            style3.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.Aqua, System.Drawing.Color.DarkSalmon);

            SLStyle style4 = sl.CreateStyle();
            style4.Border.LeftBorder.BorderStyle = BorderStyleValues.Double;
            style4.Border.LeftBorder.SetBorderThemeColor(SLThemeColorIndexValues.Accent5Color);
            style4.Border.RightBorder.BorderStyle = BorderStyleValues.Double;
            style4.Border.RightBorder.SetBorderThemeColor(SLThemeColorIndexValues.Accent5Color);

            // set column 5 with 3rd style
            sl.SetColumnStyle(5, style3);
            // set columns 7 through 9 with 4th style
            sl.SetColumnStyle(7, 9, style4);

            SLStyle style5 = sl.CreateStyle();
            style5.SetFont("Impact", 24);
            style5.Fill.SetPattern(PatternValues.LightTrellis, SLThemeColorIndexValues.Accent1Color, SLThemeColorIndexValues.Accent2Color);
            style5.Border.DiagonalBorder.BorderStyle = BorderStyleValues.DashDotDot;
            style5.Border.DiagonalBorder.SetBorderThemeColor(SLThemeColorIndexValues.Accent3Color);
            style5.Border.DiagonalUp = true;

            sl.SetCellValue(3, 4, "Do you have style?");
            sl.SetCellStyle("D3", style5);

            // set the cells from F3 to I1 to 5th style.
            // Note that this works as long as the opposite corners of the cell range is given.
            // This is effectively the same as providing
            // sl.SetCellStyle("F1", "I3", style5);
            // Or for that matter,
            // sl.SetCellStyle("I3", "F1", style5);
            sl.SetCellStyle("F3", "I1", style5);

            // this sets from rows 1 through 2, columns 11 through 13 with 5th style
            sl.SetCellStyle(1, 11, 2, 13, style5);

            // this copies the style from D3 to the range of cells A1:B2
            sl.CopyCellStyle("D3", "A1", "B2");

            // this copies the style from row 4 to row 10
            sl.CopyRowStyle(4, 10);

            // this copies the style from column 8 to columns 1 through 3
            sl.CopyColumnStyle(8, 1, 3);

            // gets the existing style from D3
            SLStyle modstyle = sl.GetCellStyle("D3");
            modstyle.Fill.SetPattern(PatternValues.DarkHorizontal, SLThemeColorIndexValues.Accent6Color, SLThemeColorIndexValues.Accent5Color);
            modstyle.RemoveBorder();

            // sets the modified style to K15
            sl.SetCellStyle("K15", modstyle);

            // the resulting spreadsheet is probably going to look ugly...

            sl.SaveAs("StyleRowColumnCell.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
