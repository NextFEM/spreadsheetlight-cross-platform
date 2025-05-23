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
            cf.SetDataBar(SLConditionalFormatDataBarValues.LightBlue);
            sl.AddConditionalFormatting(cf);

            cf = new SLConditionalFormatting("D7", "G12");
            // "false" - show data bar and value. "true" shows only the data bar
            // 20 - minimum length of data bar as percentage of cell width
            // 80 - maximum length of data bar as percentage of cell width
            // "30" - any value less than or equal to this takes the minimum length
            // "110" - any value more than or equal to this takes the maximum length
            // And use the 2nd accent color from the theme!
            cf.SetCustomDataBar(false, 20, 80,
                SLConditionalFormatMinMaxValues.Number, "30",
                SLConditionalFormatMinMaxValues.Number, "110",
                SLThemeColorIndexValues.Accent2Color);
            sl.AddConditionalFormatting(cf);

            cf = new SLConditionalFormatting("C15", "J18");
            // this shows only the data bar.
            // Note that the maximum is 150, so it can overshoot the cell
            cf.SetCustomDataBar(true, 5, 150,
                SLConditionalFormatMinMaxValues.Value, "0",
                SLConditionalFormatMinMaxValues.Percentile, "80",
                System.Drawing.Color.MediumOrchid);
            sl.AddConditionalFormatting(cf);

            sl.SaveAs("ConditionalFormatDataBar.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
