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

            // For our purposes, any date before the year 1900 is considered a historic date.
            // There's a long running battle between 1 Jan 1900 and 1 Jan 1904 being used as the
            // epoch for measuring dates in Excel. You can read all about it here:
            // http://polymathprogrammer.com/2009/10/26/the-leap-year-1900-bug-in-excel/

            // In any case, Excel turns all historic dates into text, and because it's text, it
            // stores it as a shared string, and not a number. This makes it hard to work with,
            // particularly if you want to do date manipulations.
            // Thus SpreadsheetLight offers you the chance to choose your date format before it's
            // stored into the shared string coffers. This allows you to know the format if you
            // then require to get said date back into a .NET DateTime struct.

            sl.SetCellValue(2, 2, new DateTime(1784, 1, 14), "MM/dd/yyyy");
            sl.SetCellValue(2, 5, "Ratification of Treaty of Paris");

            // I'm putting day before month because my English is graded by the British, you twat :)
            sl.SetCellValue(4, 2, new DateTime(1645, 9, 24), "dd/MM/yyyy");
            sl.SetCellValue(4, 5, "Battle of Rowton Heath");

            DateTime dtBattle = sl.GetCellValueAsDateTime(4, 2, "dd/MM/yyyy");
            sl.SetCellValue(5, 2, dtBattle.ToString("dd MMMM yyyy"));

            // If you don't set a format, the default ToString() of the DateTime struct is used.
            // It is thus considered prudent to set a date format yourself.
            sl.SetCellValue(7, 2, new DateTime(1555, 5, 29));
            sl.SetCellValue(7, 5, "Peace of Amasya");

            sl.SaveAs("HistoricDates.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}