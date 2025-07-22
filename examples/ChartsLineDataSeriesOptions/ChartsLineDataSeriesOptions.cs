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

            SLChart chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLLineChartType.StackedLine);
            chart.SetChartPosition(7, 1, 22, 8.5);

            SLDataSeriesOptions dso;

            // get the options for the 4th data series
            dso = chart.GetDataSeriesOptions(4);
            dso.Marker.Symbol = DocumentFormat.OpenXml.Drawing.Charts.MarkerStyleValues.Triangle;
            dso.Marker.Size = 10;
            // 3.25 pt
            dso.Marker.Line.Width = 3.25m;
            // 0 tint, 0% transparency
            dso.Marker.Line.SetSolidLine(SLThemeColorIndexValues.Accent5Color, 0, 0);
            // 0% transparency
            dso.Marker.Fill.SetSolidFill(System.Drawing.Color.PeachPuff, 0);
            // set the options back on the 4th data series.
            // You can also set it on another data series if you like.
            chart.SetDataSeriesOptions(4, dso);

            dso = chart.GetDataSeriesOptions(2);
            dso.Line.SetPathGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Desert);
            chart.SetDataSeriesOptions(2, dso);

            dso = chart.GetDataSeriesOptions(1);
            dso.Shadow.SetPreset(SpreadsheetLight.Drawing.SLShadowPresetValues.OuterDiagonalBottomRight);
            chart.SetDataSeriesOptions(1, dso);

            sl.InsertChart(chart);

            sl.SaveAs("ChartsLineDataSeriesOptions.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
