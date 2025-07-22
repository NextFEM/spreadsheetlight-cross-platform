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

            sl.SetCellValue(1, 1, "Age");
            sl.SetCellValue(1, 2, "Male");
            sl.SetCellValue(1, 3, "Female");

            // the negative values are important!
            // Go consult your friendly Excel user guide or search the Internet
            // for why you need them.

            sl.SetCellValue(2, 1, "<20");
            sl.SetCellValue(2, 2, -0.25);
            sl.SetCellValue(2, 3, 0.15);

            sl.SetCellValue(3, 1, "21-30");
            sl.SetCellValue(3, 2, -0.4);
            sl.SetCellValue(3, 3, 0.32);

            sl.SetCellValue(4, 1, "31-40");
            sl.SetCellValue(4, 2, -0.2);
            sl.SetCellValue(4, 3, 0.28);

            sl.SetCellValue(5, 1, "41-50");
            sl.SetCellValue(5, 2, -0.1);
            sl.SetCellValue(5, 3, 0.15);

            sl.SetCellValue(6, 1, ">50");
            sl.SetCellValue(6, 2, -0.05);
            sl.SetCellValue(6, 3, 0.1);

            sl.SetCellValue(8, 2, "=SUM(B2:B6)");
            sl.SetCellValue(8, 3, "=SUM(C2:C6)");

            SLStyle style = new SLStyle();
            style.FormatCode = "0.0%";
            // this sets the style for rows 2 to 6, columns 2 to 3
            sl.SetCellStyle(2, 2, 6, 3, style);
            sl.SetCellStyle(8, 2, style);
            sl.SetCellStyle(8, 3, style);

            SLChart chart = sl.CreateChart("A1", "C6");
            chart.SetChartType(SLBarChartType.StackedBar);
            chart.SetChartPosition(1, 4, 16, 11);
            // this makes the negative percentages display as positive too
            chart.PrimaryValueAxis.FormatCode = "0.0%;0.0%;0.0%";
            chart.PrimaryValueAxis.SourceLinked = false;
            // the maximum negative value for males is -1 (or -100% depending on how you set the values).
            // You can set it at -0.6 and it still displays fine.
            // You could get the maximum negative value from your data, but I'm too lazy to
            // do so for mine...
            chart.PrimaryValueAxis.SetOtherAxisCrossing(-1);
            chart.ShowChartTitle(false);
            chart.Title.Text = "People who watch Buffy the Vampire Slayer*";

            SLDataSeriesOptions dso = chart.GetDataSeriesOptions(1);
            // blue-ish for male. Stereotype, much?
            dso.Fill.SetSolidFill(System.Drawing.Color.RoyalBlue, 0);
            chart.SetDataSeriesOptions(1, dso);

            dso = chart.GetDataSeriesOptions(2);
            // pink-ish for female. Stereotype, much?
            dso.Fill.SetSolidFill(System.Drawing.Color.DeepPink, 0);
            chart.SetDataSeriesOptions(2, dso);

            sl.InsertChart(chart);

            sl.SetCellValue("E19", "* statistics are totally made up");

            sl.SaveAs("ChartHistogram.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
