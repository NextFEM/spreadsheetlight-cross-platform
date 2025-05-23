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

            SLStyle colstyle = sl.CreateStyle();
            colstyle.SetFontBold(true);
            colstyle.SetFontItalic(true);
            colstyle.SetPatternFill(PatternValues.Solid, SLThemeColorIndexValues.Accent4Color, SLThemeColorIndexValues.Accent4Color);

            SLStyle rowstyle1 = sl.CreateStyle();
            rowstyle1.SetPatternFill(PatternValues.Solid, SLThemeColorIndexValues.Accent2Color, SLThemeColorIndexValues.Accent2Color);

            SLStyle rowstyle2 = sl.CreateStyle();
            rowstyle2.SetGradientFill(SLGradientShadingStyleValues.Horizontal2, SLThemeColorIndexValues.Accent3Color, SLThemeColorIndexValues.Accent5Color);

            SLStyle cellstyle = sl.CreateStyle();
            cellstyle.FormatCode = "0.00";
            cellstyle.SetLeftBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Accent1Color);
            cellstyle.SetRightBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Accent1Color);
            cellstyle.SetTopBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Accent1Color);
            cellstyle.SetBottomBorder(BorderStyleValues.Thin, SLThemeColorIndexValues.Accent1Color);
            cellstyle.SetFontUnderline(UnderlineValues.Single);

            sl.SetColumnStyle(6, colstyle);
            sl.SetRowStyle(1, rowstyle1);
            sl.SetRowStyle(8, rowstyle2);
            sl.SetCellStyle("D3", "G7", cellstyle);

            Random rand = new Random();
            for (int i = 1; i <= 8; ++i)
            {
                for (int j = 1; j <= 8; ++j)
                {
                    sl.SetCellValue(i, j, rand.NextDouble() * 100.0);
                }
            }

            // copy the cell range F6:H8 to J3 (as the anchor)
            // True for cut-and-paste. The default (false) is copy-and-paste.
            sl.CopyCell("F6", "H8", "J3", true);

            // copy the cell range E1:F8 and paste at A11
            sl.CopyCell("E1", "F8", "A11");

            // copy values only (no style formatting)
            sl.CopyCell("E1", "F8", "D11", SLPasteTypeValues.Values);

            // copy only the style
            sl.CopyCell("E1", "F8", "H11", SLPasteTypeValues.Formatting);

            // transpose the cells on copy
            sl.CopyCell("E1", "F8", "G20", SLPasteTypeValues.Transpose);

            sl.AddWorksheet("Sheet2");
            // do similar copying, but this time from Sheet1 onto the new sheet.

            // notice that because we did a cut-and-paste, some of the values are missing.
            sl.CopyCellFromWorksheet("Sheet1", "E5", "H8", "I2");

            sl.CopyCellFromWorksheet("Sheet1", "E1", "F8", "A11");

            sl.CopyCellFromWorksheet("Sheet1", "E1", "F8", "D11", SLPasteTypeValues.Values);

            sl.CopyCellFromWorksheet("Sheet1", "E1", "F8", "H11", SLPasteTypeValues.Formatting);

            sl.CopyCellFromWorksheet("Sheet1", "E1", "F8", "G20", SLPasteTypeValues.Transpose);

            sl.SaveAs("CopyCell.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}