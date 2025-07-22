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
            SLThemeSettings theme = new SLThemeSettings();
            theme.ThemeName = "ColourWheel";
            theme.MajorLatinFont = "Impact";
            theme.MinorLatinFont = "Harrington";
            // this is recommended to be pure white
            theme.Light1Color = System.Drawing.Color.White;
            // this is recommended to be pure black
            theme.Dark1Color = System.Drawing.Color.Black;
            theme.Light2Color = System.Drawing.Color.GhostWhite;
            theme.Dark2Color = System.Drawing.Color.DarkSlateGray;
            theme.Accent1Color = System.Drawing.Color.Red;
            theme.Accent2Color = System.Drawing.Color.OrangeRed;
            theme.Accent3Color = System.Drawing.Color.Yellow;
            theme.Accent4Color = System.Drawing.Color.LawnGreen;
            theme.Accent5Color = System.Drawing.Color.DeepSkyBlue;
            theme.Accent6Color = System.Drawing.Color.DarkViolet;
            theme.Hyperlink = System.Drawing.Color.Blue;
            theme.FollowedHyperlinkColor = System.Drawing.Color.Purple;

            SLDocument sl = new SLDocument(theme);

            SLStyle style;

            style = sl.CreateStyle();
            style.SetFont(FontSchemeValues.Major, 11);
            sl.SetCellStyle(2, 7, style);
            sl.SetCellValue(2, 7, "I'm Major Fontsalot.");

            // we can reuse this way because at this point, we've only used the font part.
            // Of course, by default, the minor font is used. This is to show you how to
            // explicitly set a minor font.
            style.SetFont(FontSchemeValues.Minor, 11);
            sl.SetCellStyle(3, 7, style);
            sl.SetCellValue(3, 7, "And I'm his lackey. And a minor to boot. Is there a union against this or something?");

            sl.SetCellValue(5, 8, "Accent 1 colour");
            sl.SetCellValue(6, 8, "Accent 2 colour");
            sl.SetCellValue(7, 8, "Accent 3 colour");
            sl.SetCellValue(8, 8, "Accent 4 colour");
            sl.SetCellValue(9, 8, "Accent 5 colour");
            sl.SetCellValue(10, 8, "Accent 6 colour");

            style = sl.CreateStyle();
            style.Fill.SetPatternType(PatternValues.Solid);

            // we're going to reuse the style class because we're just changing the colour.
            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent1Color);
            sl.SetCellStyle(5, 7, style);

            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent2Color);
            sl.SetCellStyle(6, 7, style);

            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent3Color);
            sl.SetCellStyle(7, 7, style);

            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent4Color);
            sl.SetCellStyle(8, 7, style);

            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent5Color);
            sl.SetCellStyle(9, 7, style);

            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent6Color);
            sl.SetCellStyle(10, 7, style);

            sl.SaveAs("ThemeNew.xlsx");


            // Or you can use a built-in theme
            sl = new SLDocument(SLThemeTypeValues.Metro);

            // note that the following chunk of code is exactly the same as above
            // The only difference is that the theme used is different.

            style = sl.CreateStyle();
            style.SetFont(FontSchemeValues.Major, 11);
            sl.SetCellStyle(2, 7, style);
            sl.SetCellValue(2, 7, "I'm Major Fontsalot.");

            // we can reuse this way because at this point, we've only used the font part.
            // Of course, by default, the minor font is used. This is to show you how to
            // explicitly set a minor font.
            style.SetFont(FontSchemeValues.Minor, 11);
            sl.SetCellStyle(3, 7, style);
            sl.SetCellValue(3, 7, "And I'm his lackey. And a minor to boot. Is there a union against this or something?");

            sl.SetCellValue(5, 8, "Accent 1 colour");
            sl.SetCellValue(6, 8, "Accent 2 colour");
            sl.SetCellValue(7, 8, "Accent 3 colour");
            sl.SetCellValue(8, 8, "Accent 4 colour");
            sl.SetCellValue(9, 8, "Accent 5 colour");
            sl.SetCellValue(10, 8, "Accent 6 colour");

            style = sl.CreateStyle();
            style.Fill.SetPatternType(PatternValues.Solid);

            // we're going to reuse the style class because we're just changing the colour.
            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent1Color);
            sl.SetCellStyle(5, 7, style);

            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent2Color);
            sl.SetCellStyle(6, 7, style);

            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent3Color);
            sl.SetCellStyle(7, 7, style);

            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent4Color);
            sl.SetCellStyle(8, 7, style);

            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent5Color);
            sl.SetCellStyle(9, 7, style);

            style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent6Color);
            sl.SetCellStyle(10, 7, style);

            sl.SaveAs("ThemeBuiltin.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
