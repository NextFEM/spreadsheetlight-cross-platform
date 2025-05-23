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
            chart.SetChartType(SLColumnChartType.Column3D);
            chart.SetChartPosition(7, 1, 22, 8.5);

            // 10 pt width!
            chart.Floor.Border.Width = 10;
            // 0 tint, 0 transparency
            chart.Floor.Border.SetSolidLine(SLThemeColorIndexValues.Accent6Color, 0, 0);
            // 0 transparency
            chart.Floor.Fill.SetSolidFill(System.Drawing.Color.BlanchedAlmond, 0);

            // This stretches the picture to fill up the space
            // 0 for left offset, right offset, top offset, bottom offset and transparency
            chart.SideWall.Fill.SetPictureFill("mandelbrot.png", 0, 0, 0, 0, 0);
            // Use the other SetPictureFill() overload to tile the picture.

            // Rainbow colours!
            chart.BackWall.Fill.SetPathGradient(SLGradientPresetValues.Rainbow);
            chart.BackWall.Border.SetSolidLine(System.Drawing.Color.Aquamarine, 0);
            chart.BackWall.Border.Width = 5;

            sl.InsertChart(chart);

            sl.SaveAs("Chart3DFloorSideBack.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
