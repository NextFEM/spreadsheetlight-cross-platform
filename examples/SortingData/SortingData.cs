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

            int i, j;
            for (i = 1; i <= 20; ++i)
            {
                for (j = 1; j <= 10; ++j)
                {
                    sl.SetCellValue(i, j, i * j);
                }
            }

            // sort cell range A1:E7 by the column D in descending order
            sl.Sort("A1", "E7", "D", false);

            // sort cell range with rows from 5 through 18, columns 7 through 9
            // (or G5:I18)
            // "false" - sort by row (instead of by column)
            // sorted on index 11, so this is row index 11
            // "false" - sort in descending order
            sl.Sort(5, 7, 18, 9, false, 11, false);

            sl.SaveAs("SortingData.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
