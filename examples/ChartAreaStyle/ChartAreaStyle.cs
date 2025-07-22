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

            // 25% transparency
            chart.Fill.SetSolidFill(System.Drawing.Color.AntiqueWhite, 25);

            // 5 pt
            chart.Border.Width = 5;
            // 0 tint, 0 transparency
            chart.Border.SetSolidLine(SLThemeColorIndexValues.Accent5Color, 0, 0);

            // 6 pt width, 6 pt height
            chart.Format3D.SetBevelTop(DocumentFormat.OpenXml.Drawing.BevelPresetValues.Angle, 6, 6);

            chart.Shadow.SetPreset(SLShadowPresetValues.PerspectiveDiagonalUpperRight);

            sl.InsertChart(chart);

            sl.SaveAs("ChartAreaStyle.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
