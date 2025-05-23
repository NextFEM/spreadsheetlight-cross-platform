using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using SpreadsheetLight.Drawing;

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

            SLChart chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLColumnChartType.ClusteredColumn);
            chart.SetChartPosition(7, 1, 22, 8.5);

            chart.PrimaryTextAxis.ShowTitle = true;
            chart.PrimaryTextAxis.Title.Text = "Primary Axis Title";
            // 0 degree means from left to right
            chart.PrimaryTextAxis.Title.Border.SetLinearGradient(SLGradientPresetValues.Rainbow, 0);
            // 3 pt width, the better to see the rainbow
            chart.PrimaryTextAxis.Title.Border.Width = 3;

            SLFont font;

            SLRstType rst = new SLRstType();
            font = new SLFont();
            font.Italic = true;
            font.Bold = false;
            font.SetFont("Impact", 16);
            // 60% lightening
            font.SetFontThemeColor(SLThemeColorIndexValues.Accent1Color, 0.6);
            rst.AppendText("Sec ", font);

            font = new SLFont();
            font.Bold = false;
            font.Underline = UnderlineValues.Single;
            // mmm... salmon...
            font.FontColor = System.Drawing.Color.LightSalmon;
            rst.AppendText("Axis Title", font);

            chart.PrimaryValueAxis.ShowTitle = true;
            chart.PrimaryValueAxis.Title.SetTitle(rst);
            // 0 degree rotation
            chart.PrimaryValueAxis.Title.SetHorizontalTextDirection(SLTextVerticalAlignment.MiddleCentered, 0);
            // 25% darkening, 0% transparency
            chart.PrimaryValueAxis.Title.Fill.SetSolidFill(SLThemeColorIndexValues.Accent3Color, -0.25, 0);

            sl.InsertChart(chart);

            sl.SaveAs("ChartAxisTitle.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
