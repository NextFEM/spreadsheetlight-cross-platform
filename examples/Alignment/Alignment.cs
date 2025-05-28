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

            sl.SetCellValue("B10", "This is a short sentence.");

            SLStyle style = sl.CreateStyle();
            style.Alignment.Horizontal = HorizontalAlignmentValues.Left;
            // Alternatively, use the shortcut function:
            // style.SetHorizontalAlignment(HorizontalAlignmentValues.Left);

            // each indent is 3 spaces, so this is 15 spaces total
            style.Alignment.Indent = 5;
            style.Alignment.JustifyLastLine = true;
            style.Alignment.ReadingOrder = SLAlignmentReadingOrderValues.RightToLeft;
            style.Alignment.ShrinkToFit = true;
            style.Alignment.TextRotation = 30;

            style.SetVerticalAlignment(VerticalAlignmentValues.Center);
            // Alternatively, use the "long-form" version:
            // style.Alignment.Vertical = VerticalAlignmentValues.Center;

            style.SetWrapText(true);
            // Alternatively, use the "long-form" version:
            // style.Alignment.WrapText = true;

            // it will look weird... hey that's a lot of alignment settings!
            sl.SetCellStyle(10, 2, style);

            sl.SaveAs("Alignment.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
