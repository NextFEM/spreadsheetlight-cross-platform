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

            chart.Legend.LegendPosition = DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues.TopRight;
            chart.Legend.Fill.SetRadialGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Gold, SpreadsheetLight.Drawing.SLGradientDirectionValues.Center);
            // 0% transparency
            chart.Legend.Border.SetSolidLine(System.Drawing.Color.Orange, 0);
            chart.Legend.Shadow.SetPreset(SpreadsheetLight.Drawing.SLShadowPresetValues.PerspectiveDiagonalUpperLeft);

            sl.InsertChart(chart);

            sl.SaveAs("ChartLegendExample.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
