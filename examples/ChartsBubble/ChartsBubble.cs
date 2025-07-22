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

            // I don't understand how bubble charts are used...
            // You might know the correct cell content to use...
            sl.SetCellValue("C2", "Col1");
            sl.SetCellValue("D2", "Col2");
            sl.SetCellValue("E2", "Size");

            Random rand = new Random();
            for (int i = 3; i <= 6; ++i)
            {
                sl.SetCellValue(i, 3, 9000 * rand.NextDouble() + 1000);
                sl.SetCellValue(i, 4, 9000 * rand.NextDouble() + 1000);
                sl.SetCellValue(i, 5, 50 * rand.NextDouble() + 20);
            }

            SLChart chart = sl.CreateChart("C2", "E6");
            chart.SetChartType(SLBubbleChartType.Bubble3D);
            chart.SetChartPosition(7, 1, 22, 8.5);

            sl.InsertChart(chart);

            sl.SaveAs("ChartsBubble.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
