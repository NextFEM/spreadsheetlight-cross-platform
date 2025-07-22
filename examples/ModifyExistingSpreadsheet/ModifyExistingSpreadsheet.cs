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
            // SpreadsheetLight works on the idea of a currently selected worksheet.
            // If no worksheet name is provided on opening an existing spreadsheet,
            // the first available worksheet is selected.
            SLDocument sl = new SLDocument("ModifyExistingSpreadsheetOriginal.xlsx", "Sheet2");

            sl.SetCellValue("E6", "Let's party!!!!111!!!1");

            sl.SelectWorksheet("Sheet3");
            sl.SetCellValue("E6", "Before anyone calls the popo!");

            sl.AddWorksheet("DanceFloor");
            sl.SetCellValue("B4", "Who let the dogs out?");
            sl.SetCellValue("B5", "Woof!");

            sl.SaveAs("ModifyExistingSpreadsheetModified.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
