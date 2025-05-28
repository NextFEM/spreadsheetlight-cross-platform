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

            // hide row 3
            sl.HideRow(3);

            // hide columns 5 through 8
            sl.HideColumn(5, 8);

            sl.SaveAs("HideRowColumn.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
