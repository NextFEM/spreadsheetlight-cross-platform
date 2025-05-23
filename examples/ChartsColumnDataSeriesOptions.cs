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
            chart.SetChartType(SLColumnChartType.ClusteredColumn3D);
            chart.SetChartPosition(7, 1, 22, 8.5);

            SLDataSeriesOptions dso;
            // get the options from the 2nd data series
            dso = chart.GetDataSeriesOptions(2);
            dso.Shape = DocumentFormat.OpenXml.Drawing.Charts.ShapeValues.Pyramid;
            // 10% transparency
            dso.Fill.SetSolidFill(System.Drawing.Color.MediumOrchid, 10);
            // Set on the 2nd data series.
            // Make sure you set the options on the correct data series index that
            // you got it from.
            // Or not, depending on what you want to achieve...
            chart.SetDataSeriesOptions(2, dso);

            dso = chart.GetDataSeriesOptions(4);
            // in this case, the shadow is almost imperceptible. Just look harder, ok?
            dso.Shadow.SetPreset(SpreadsheetLight.Drawing.SLShadowPresetValues.PerspectiveDiagonalUpperRight);
            // 0% transparency
            dso.Line.SetSolidLine(System.Drawing.Color.Orange, 0);
            chart.SetDataSeriesOptions(4, dso);

            sl.InsertChart(chart);

            sl.SaveAs("ChartsColumnDataSeriesOptions.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
