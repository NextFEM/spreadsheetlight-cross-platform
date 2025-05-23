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

            sl.SetCellValue(5, 5, "Make a pattern");

            SLStyle style = sl.CreateStyle();
            // solid pattern, foreground is red, background is blue
            // But it's more complicated than this. See explanation below...
            style.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.Red, System.Drawing.Color.Blue);
            sl.SetCellStyle(6, 5, style);

            // The typical use of fill is a solid fill with a colour.
            // Internally, Excel uses a solid fill and wait for it... the foreground colour
            // property to store it.
            // The background colour property only comes into play when the fill type
            // is not a solid fill, say a hatch pattern.
            // So for typical use of solid fills, you can just set the pattern type
            // (with the solid enum value) and the foreground colour.
            // Like so:
            //style.Fill.SetPatternType(PatternValues.Solid);
            //style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent1Color);

            // There are shortcut functions for setting normal pattern and gradient fills.

            style = sl.CreateStyle();
            // dark trellis pattern, foreground is accent 2 colour, background is accent 5 colour
            style.SetPatternFill(PatternValues.DarkTrellis, SLThemeColorIndexValues.Accent2Color, SLThemeColorIndexValues.Accent5Color);
            // Alternatively, use the "long-form" version:
            // style.Fill.SetPattern(PatternValues.DarkTrellis, SLThemeColorIndexValues.Accent2Color, SLThemeColorIndexValues.Accent5Color);
            sl.SetCellStyle(8, 5, style);

            style = sl.CreateStyle();
            // The SLGradientShadingStyleValues enumeration follows Excel gradient options.
            // DiagonalUp3 means interpolate from top-left to bottom-right,
            // using the 1st colour at the top-left, the 2nd color in the middle, and
            // then use the 1st colour at the bottom-right.
            // In this case, the 1st colour is accent 1 colour, and the
            // 2nd colour is the accent 2 colour.
            style.SetGradientFill(SLGradientShadingStyleValues.DiagonalUp3, SLThemeColorIndexValues.Accent1Color, SLThemeColorIndexValues.Accent2Color);
            // Alternatively, use the "long-form" version:
            // style.Fill.SetGradient(SLGradientShadingStyleValues.DiagonalUp3, SLThemeColorIndexValues.Accent1Color, SLThemeColorIndexValues.Accent2Color);
            sl.SetCellStyle(10, 5, style);

            // now for some really fancy gradients...

            style = sl.CreateStyle();
            // set linear interpolation, horizontal interpolation (0),
            // start from left (0) to right (1),
            // don't care about top (null) and bottom (null)
            style.Fill.SetCustomGradient(GradientValues.Linear, 0, 0, 1, null, null);
            // interpolation starts with accent 1 colour
            style.Fill.AppendGradientStop(0.0, SLThemeColorIndexValues.Accent1Color);
            // at midpoint (0.5), set CadetBlue as colour
            style.Fill.AppendGradientStop(0.5, System.Drawing.Color.CadetBlue);
            // at 80% of the way, set accent 5 colour
            style.Fill.AppendGradientStop(0.8, SLThemeColorIndexValues.Accent5Color);
            // at the end, set accent 2 colour, darkened 40%
            style.Fill.AppendGradientStop(1.0, SLThemeColorIndexValues.Accent2Color, -0.4);
            sl.SetCellStyle(12, 5, style);

            sl.SaveAs("Patterns.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
