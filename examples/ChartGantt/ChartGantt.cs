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

            // WARNING: I have zero experience in home making stuff.

            sl.SetColumnWidth(1, 25);

            sl.SetCellValue(1, 2, "Start Date");
            sl.SetCellValue(1, 3, "Days");

            sl.SetCellValue(2, 1, "Do landscaping");
            sl.SetCellValue(2, 2, new DateTime(2012, 11, 24));
            sl.SetCellValue(2, 3, 2);

            sl.SetCellValue(3, 1, "Change to wooden flooring");
            sl.SetCellValue(3, 2, new DateTime(2012, 11, 26));
            sl.SetCellValue(3, 3, 10);

            sl.SetCellValue(4, 1, "Wax the floor");
            sl.SetCellValue(4, 2, new DateTime(2012, 12, 10));
            sl.SetCellValue(4, 3, 5);

            sl.SetCellValue(5, 1, "Paint walls");
            sl.SetCellValue(5, 2, new DateTime(2012, 12, 15));
            sl.SetCellValue(5, 3, 4);

            sl.SetCellValue(6, 1, "Move furniture in");
            sl.SetCellValue(6, 2, new DateTime(2012, 12, 20));
            sl.SetCellValue(6, 3, 3);

            sl.SetCellValue(7, 1, "Rest");
            sl.SetCellValue(7, 2, new DateTime(2012, 12, 23));
            sl.SetCellValue(7, 3, 1);

            sl.SetCellValue(7, 5, "Yeah we survived! Suck it Mayan long count calendar!");

            sl.SetCellValue(8, 1, "Holiday party!");
            sl.SetCellValue(8, 2, new DateTime(2012, 12, 25));
            sl.SetCellValue(8, 3, 1);

            SLStyle style = new SLStyle();
            style.FormatCode = "d-mmm";
            sl.SetCellStyle(2, 2, 8, 2, style);

            SLChart chart = sl.CreateChart("A1", "C8");
            chart.SetChartType(SLBarChartType.StackedBar);
            chart.SetChartPosition(10, 1, 24, 11);

            chart.HideChartLegend();
            chart.PrimaryTextAxis.InReverseOrder = true;
            chart.PrimaryTextAxis.SetMaximumOtherAxisCrossing();
            // it's not exactly 1 Jan 1900 because there's the incorrect 29 Feb 1900
            // but it's not worth quibbling about...
            chart.PrimaryValueAxis.Minimum = (new DateTime(2012, 11, 24) - new DateTime(1900, 1, 1)).Days;
            // we add more days to the last day so the last day is also included
            // Go experiment with the values...
            chart.PrimaryValueAxis.Maximum = (new DateTime(2012, 12, 30) - new DateTime(1900, 1, 1)).Days;
            // 7 days. Set the interval as weekly
            chart.PrimaryValueAxis.MajorUnit = 7;

            SLDataSeriesOptions dso = chart.GetDataSeriesOptions(1);
            dso.Fill.SetNoFill();
            dso.Line.SetNoLine();
            chart.SetDataSeriesOptions(1, dso);

            sl.InsertChart(chart);

            sl.SaveAs("ChartGantt.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
