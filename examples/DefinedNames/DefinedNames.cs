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

            for (int i = 1; i < 6; ++i)
            {
                for (int j = 1; j < 6; ++j)
                {
                    sl.SetCellValue(i, j, i * j);
                }
            }

            // Defined names are what were used to be known as named cells
            // or named ranges.

            sl.SetDefinedName("MySum", "Sheet1!$B$2:$D$4");
            sl.SetDefinedName("MyValue", "Sheet1!$D$5");

            sl.SetCellValue(8, 2, "=SUM(MySum)");
            sl.SetCellValue(9, 2, "=MyValue");

            sl.SaveAs("DefinedNames.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
