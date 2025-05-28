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

            for (int i = 1; i < 20; ++i)
            {
                for (int j = 1; j < 15; ++j)
                {
                    sl.SetCellValue(i, j, string.Format("R{0}C{1}", i, j));
                }
            }

            // set rows 3 through 8 with a row height of 40 points.
            // The row height is typically around 15 points for most fonts.
            sl.SetRowHeight(3, 8, 40);

            // set column 6 with a column width of 15.
            // Column width is measured as the number of characters of maximum digit
            // width of digits 0-9 as rendered in the Normal font.
            // Confused? The width is around 8.8 for most fonts.
            sl.SetColumnWidth(6, 15);

            sl.SaveAs("SetRowColumnDimension.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
