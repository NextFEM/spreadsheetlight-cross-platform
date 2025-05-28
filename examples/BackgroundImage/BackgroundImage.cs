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
            SLDocument sl = new SLDocument("BackgroundImageOriginal.xlsx", "Sheet1");

            // this will overwrite any existing background image
            sl.AddBackgroundPicture("julia.png");
            sl.SetCellValue(2, 2, "The Art Maestro was here! Muahahaha!");

            sl.SelectWorksheet("Sheet2");
            // this deletes any existing background image
            sl.DeleteBackgroundPicture();
            sl.SetCellValue(2, 2, "You're not worthy of having this art piece!");

            sl.SelectWorksheet("Sheet3");
            // just to show that background images are untouched if no operations are done
            sl.SetCellValue(2, 2, "Hmmm... this doesn't seem to be worth a lot...");

            sl.AddWorksheet("Sheet4");
            sl.AddBackgroundPicture("mandelbrot.png");
            sl.SetCellValue(2, 2, "Here's a gift from the Art Maestro!");

            sl.SaveAs("BackgroundImageModified.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
