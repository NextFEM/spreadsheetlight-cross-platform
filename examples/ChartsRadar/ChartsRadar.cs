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

            double fChartHeight = 15.0;
            double fChartWidth = 7.5;

            SLChart chart;

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLRadarChartType.Radar);
            chart.SetChartPosition(1, 9, 1 + fChartHeight, 9 + fChartWidth);
            sl.InsertChart(chart);

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLRadarChartType.RadarWithMarkers);
            chart.SetChartStyle(SLChartStyle.Style10);
            chart.SetChartPosition(7, 1, 7 + fChartHeight, 1 + fChartWidth);
            sl.InsertChart(chart);

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLRadarChartType.FilledRadar);
            chart.SetChartStyle(SLChartStyle.Style18);
            chart.SetChartPosition(16, 9, 16 + fChartHeight, 9 + fChartWidth);
            sl.InsertChart(chart);

            sl.SaveAs("ChartsRadar.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
