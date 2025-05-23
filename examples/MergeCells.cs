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

            sl.SetCellValue("B2", "This is a merged cell");

            // merge all cells in the cell range B2:G8
            sl.MergeWorksheetCells("B2", "G8");

            // merge all cells from rows 10 through 12, columns 4 through 6
            // This is basically the cell range D10:F12
            sl.MergeWorksheetCells(10, 4, 12, 6);

            // merge alls cells from rows 15 through 4, columns 12 through 9
            // Note that the order of the corresponding row and column indices
            // doesn't matter. This is the cell range I4:L15
            sl.MergeWorksheetCells(15, 12, 4, 9);

            // unmerge the cell range D10:F12
            sl.UnmergeWorksheetCells("D10", "F12");

            sl.SaveAs("MergeCells.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
