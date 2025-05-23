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
                    sl.SetCellValue(i, j, 400 * rand.NextDouble() - 200.0);
                }
            }

            SLDataBarOptions dbo;
            SLConditionalFormatting cf;

            // the default is true, for defaulting to Excel 2010 specific data bars.
            // This is in case you still want to generate a data bar that's more for Excel 2007.
            // Just set false then.
            dbo = new SLDataBarOptions(SLConditionalFormatDataBarValues.Green, true);
            cf = new SLConditionalFormatting("B2", "F7");
            cf.SetCustomDataBar(dbo);
            sl.AddConditionalFormatting(cf);

            dbo = new SLDataBarOptions(SLConditionalFormatDataBarValues.Orange);
            // this essentially ignores the original color type you just set... you don't like orange, huh?
            dbo.FillColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
            dbo.NegativeFillColor.Color = System.Drawing.Color.IndianRed;
            dbo.Border = true;
            dbo.BorderColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
            cf = new SLConditionalFormatting("H4", "J9");
            cf.SetCustomDataBar(dbo);
            sl.AddConditionalFormatting(cf);

            dbo = new SLDataBarOptions(SLConditionalFormatDataBarValues.Purple);
            dbo.MinimumType = SLConditionalFormatAutoMinMaxValues.Number;
            dbo.MinimumValue = "-50";
            dbo.MaximumType = SLConditionalFormatAutoMinMaxValues.Number;
            dbo.MaximumValue = "150";
            dbo.AxisColor.Color = System.Drawing.Color.DodgerBlue;
            dbo.AxisPosition = DocumentFormat.OpenXml.Office2010.Excel.DataBarAxisPositionValues.Middle;
            cf = new SLConditionalFormatting("C12", "G19");
            cf.SetCustomDataBar(dbo);
            sl.AddConditionalFormatting(cf);

            sl.SaveAs("ConditionalFormatDataBar2010.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
