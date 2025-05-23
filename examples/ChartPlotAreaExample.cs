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

            // 45 degrees, meaning interpolate from top-left to bottom-right
            chart.PlotArea.Fill.SetLinearGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Silver, 45);

            // 0 degrees, meaning interpolate from left to right
            chart.PlotArea.Border.SetLinearGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Peacock, 0);

            // 3 pt for width and height
            chart.PlotArea.Format3D.SetBevelTop(DocumentFormat.OpenXml.Drawing.BevelPresetValues.Divot, 3, 3);
            // 4 pt for width and height
            chart.PlotArea.Format3D.SetBevelBottom(DocumentFormat.OpenXml.Drawing.BevelPresetValues.Cross, 4, 4);
            chart.PlotArea.Format3D.Material = DocumentFormat.OpenXml.Drawing.PresetMaterialTypeValues.DarkEdge;
            chart.PlotArea.Format3D.Lighting = DocumentFormat.OpenXml.Drawing.LightRigValues.Harsh;

            chart.PlotArea.Shadow.SetPreset(SpreadsheetLight.Drawing.SLShadowPresetValues.PerspectiveDiagonalUpperRight);

            sl.InsertChart(chart);

            sl.SaveAs("ChartPlotAreaExample.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
