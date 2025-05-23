using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            SLDocument sl = new SLDocument();

            Random rand = new Random();
            int i, j;
            for (i = 1; i <= 20; ++i)
            {
                for (j = 1; j <= 10; ++j)
                {
                    sl.SetCellValue(i, j, 200 * rand.NextDouble());
                }
            }

            SLConditionalFormatting cf;

            cf = new SLConditionalFormatting("B2", "H5");
            cf.SetColorScale(SLConditionalFormatColorScaleValues.RedYellowGreen);
            sl.AddConditionalFormatting(cf);

            cf = new SLConditionalFormatting("D7", "G12");
            // the minimum color is GreenYellow for values at 20% or below
            // (so it's <= 40 because our cell values range from 0 to 200)
            // the maximum color is OrangeRed for values at 80% or above
            // (so it's >= 160 because our cell values range from 0 to 200)
            cf.SetCustom2ColorScale(SLConditionalFormatMinMaxValues.Percent, "20", System.Drawing.Color.GreenYellow,
                SLConditionalFormatMinMaxValues.Percent, "80", System.Drawing.Color.OrangeRed);
            sl.AddConditionalFormatting(cf);

            cf = new SLConditionalFormatting("C15", "J18");
            // the minimum is colored with accent 1 color that's lightened 20%
            // the midpoint is at the 35th percentile, colored with accent 3 color that's darkened 10%
            // the maximum is colored with accent 6 color that's lightened 50%
            cf.SetCustom3ColorScale(SLConditionalFormatMinMaxValues.Value, "0", SLThemeColorIndexValues.Accent1Color, 0.2,
                SLConditionalFormatRangeValues.Percentile, "35", SLThemeColorIndexValues.Accent3Color, -0.1,
                SLConditionalFormatMinMaxValues.Value, "0", SLThemeColorIndexValues.Accent6Color, 0.5);
            sl.AddConditionalFormatting(cf);

            sl.SaveAs("ConditionalFormatColorScale.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
