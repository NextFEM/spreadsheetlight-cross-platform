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
            chart.SetChartType(SLColumnChartType.ClusteredColumn);
            chart.SetChartPosition(7, 1, 22, 8.5);

            // this plots the 3rd data series as a normal line
            // (as opposed to the stacked and 100% stacked versions) with markers
            // on the primary axis.
            chart.PlotDataSeriesAsPrimaryLineChart(3, SLChartDataDisplayType.Normal, true);

            // this plots the 4th data series as a normal line without the markers
            // on the secondary axis.
            chart.PlotDataSeriesAsSecondaryLineChart(4, SLChartDataDisplayType.Normal, false);

            // this plots the 2nd data series as a normal area chart on the primary axis.
            chart.PlotDataSeriesAsPrimaryAreaChart(2, SLChartDataDisplayType.Normal);

            // Combination charts are complicated. Only weak validation checks are done.
            // Just because you can plot the data series as whatever chart types and on
            // whichever axis, doesn't mean Excel will render them. Or if the final
            // combination chart is valid (or makes sense). Proceed with caution.

            // However, if Excel allows a particular combination chart, you should be able
            // to do it just fine.

            sl.InsertChart(chart);

            sl.SaveAs("ChartsColumnLineAreaCombination.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
