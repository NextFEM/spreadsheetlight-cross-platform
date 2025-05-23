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

            sl.SetCellValue("C2", "Apple");
            sl.SetCellValue("D2", "Banana");
            sl.SetCellValue("E2", "Cherry");
            sl.SetCellValue("F2", "Durian");
            sl.SetCellValue("G2", "Elderberry");
            sl.SetCellValue("B3", "North");
            sl.SetCellValue("B4", "South");
            sl.SetCellValue("B5", "East");
            sl.SetCellValue("B6", "West");

            Random rand = new Random();
            for (int i = 3; i <= 6; ++i)
            {
                for (int j = 3; j <= 7; ++j)
                {
                    sl.SetCellValue(i, j, 9000 * rand.NextDouble() + 1000);
                }
            }

            SLChart chart;

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLBarChartType.ClusteredBar);
            chart.SetChartPosition(7, 1, 22, 8.5);

            SLFont ft;
            SLRstType rst = sl.CreateRstType();

            ft = sl.CreateFont();
            ft.SetFont("Impact", 24);
            rst.AppendText("Example ", ft);

            ft = sl.CreateFont();
            ft.SetFont(FontSchemeValues.Major, 16);
            ft.Italic = true;
            ft.Underline = UnderlineValues.Single;
            rst.AppendText("Chart ", ft);

            ft = sl.CreateFont();
            ft.SetFontThemeColor(SLThemeColorIndexValues.Accent1Color);
            rst.AppendText("Title", ft);

            chart.Title.SetTitle(rst);
            // set true for title to overlap the plot area
            chart.ShowChartTitle(false);

            // use accent 5 color, at 0.8 tint, and at 50% transparency
            chart.Title.Fill.SetSolidFill(SLThemeColorIndexValues.Accent5Color, 0.8, 50);

            // 0% transparency
            chart.Title.Border.SetSolidLine(System.Drawing.Color.MediumPurple, 0);

            chart.Title.Shadow.SetPreset(SpreadsheetLight.Drawing.SLShadowPresetValues.PerspectiveDiagonalUpperLeft);

            // 6 pt for width and height
            chart.Title.Format3D.SetBevelTop(DocumentFormat.OpenXml.Drawing.BevelPresetValues.RelaxedInset, 6, 6);
            chart.Title.Format3D.Material = DocumentFormat.OpenXml.Drawing.PresetMaterialTypeValues.Metal;
            chart.Title.Format3D.Lighting = DocumentFormat.OpenXml.Drawing.LightRigValues.Freezing;

            // true for left-to-right, false for right-to-left
            chart.Title.SetStackedTextDirection(SpreadsheetLight.Drawing.SLTextHorizontalAlignment.CenterMiddle, true);

            sl.InsertChart(chart);

            sl.SaveAs("ChartTitleExample.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
