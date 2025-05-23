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
            // "false" - not between (instead of true for "between")
            // This highlights cells fulfilling: value < 50 OR 150 < value
            cf.HighlightCellsBetween(false, "50", "150", SLHighlightCellsStyleValues.RedBorder);
            sl.AddConditionalFormatting(cf);

            cf = new SLConditionalFormatting("D7", "G12");
            // false - in bottom range (instead of true for in top range)
            // 25 for rank
            // true - is percentage (instead of false for number of items)
            // This means it's bottom 25% (instead of bottom 25 items)
            cf.HighlightCellsInTopRange(false, 25, true, SLHighlightCellsStyleValues.LightRedFill);
            sl.AddConditionalFormatting(cf);

            // clear cell values in cell range D16:E17
            sl.ClearCellContent("D16", "E17");

            // this will create an error cell because there's no function called "SpreadsheetLight"
            // nor is there a defined name called "SpreadsheetLight"
            sl.SetCellValue("F18", "=SpreadsheetLight");

            // this will create duplicate values
            sl.SetCellValue("H16", "asdf");
            sl.SetCellValue("H17", "asdf");

            cf = new SLConditionalFormatting("C15", "J18");
            // "true" - contain blanks ("false" for "not containing blanks")
            cf.HighlightCellsContainingBlanks(true, SLHighlightCellsStyleValues.YellowFillWithDarkYellowText);
            sl.AddConditionalFormatting(cf);

            cf = new SLConditionalFormatting("C15", "J18");
            // "true" - contain errors ("false" for "not containing errors")
            cf.HighlightCellsContainingErrors(true, SLHighlightCellsStyleValues.LightRedFillWithDarkRedText);
            sl.AddConditionalFormatting(cf);

            cf = new SLConditionalFormatting("C15", "J18");
            cf.HighlightCellsWithDuplicates(SLHighlightCellsStyleValues.GreenFillWithDarkGreenText);
            sl.AddConditionalFormatting(cf);

            sl.SaveAs("ConditionalFormatHighlightCells.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
