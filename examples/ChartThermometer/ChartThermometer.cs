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

            sl.SetCellValue(1, 2, "Sales");
            for (int i = 1; i <= 14; ++i)
            {
                sl.SetCellValue(i + 1, 1, string.Format("Day {0}", i));
            }

            Random rand = new Random();
            for (int i = 2; i <= 9; ++i)
            {
                sl.SetCellValue(i, 2, rand.Next(1, 9));
            }

            sl.SetCellValue(17, 1, "Total");
            sl.SetCellValue(17, 2, "=SUM(B2:B15)");

            sl.SetCellValue(18, 1, "Goal");
            sl.SetCellValue(18, 2, 75);

            // this is the cell we use to create the chart.
            sl.SetCellValue(20, 2, "=B17/B18");

            SLStyle style = sl.CreateStyle();
            style.FormatCode = "0%";
            sl.SetCellStyle(20, 2, style);

            SLBarChartOptions bcoptions = new SLBarChartOptions();
            bcoptions.Overlap = 0;
            bcoptions.GapWidth = 0;

            SLChart chart = sl.CreateChart("A19", "B20");
            chart.SetChartType(SLColumnChartType.ClusteredColumn, bcoptions);
            chart.SetChartPosition(1.5, 2.5, 16.5, 5.5);
            chart.HideChartLegend();
            chart.HidePrimaryTextAxis();
            // this is with respect to the cell we use to create the chart.
            // The value ranges from 0 to 1 (in percentage).
            chart.PrimaryValueAxis.Minimum = 0;
            chart.PrimaryValueAxis.Maximum = 1;
            chart.PrimaryValueAxis.ShowMajorGridlines = false;
            chart.PlotArea.Fill.SetLinearGradient(SLGradientPresetValues.Silver, 0);

            SLDataPointOptions dcoptions = chart.CreateDataPointOptions();
            // 90 degrees, so it's top to bottom
            dcoptions.Fill.SetLinearGradient(SLGradientPresetValues.Fire, 90);
            chart.SetDataPointOptions(1, 1, dcoptions);

            sl.InsertChart(chart);

            sl.SaveAs("ChartThermometer.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
