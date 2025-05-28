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

            // Named cell styles include:
            // Normal, Bad, Good, Neutral
            // Calculation, Check Cell, Explanatory Text, Input
            // Linked Cell, Note, Output, Warning Text
            // Heading 1, Heading 2, Heading 3, Heading 4, Title, Total
            // Accents 1 through 6, with variations of 20%, 40% and 60% lightening
            // Comma, Comma [0], Currency, Currency [0], Percent

            sl.ApplyNamedCellStyle("B1", "B8", SLNamedCellStyleValues.WarningText);
            for (int i = 1; i < 9; ++i) sl.SetCellValue(i, 2, "Warning " + i.ToString());

            SLStyle style = sl.CreateStyle();
            style.ApplyNamedCellStyle(SLNamedCellStyleValues.Heading1);
            style.Border.TopBorder.BorderStyle = BorderStyleValues.Thick;
            style.Border.TopBorder.Color = System.Drawing.Color.HotPink;
            sl.SetCellStyle("D4", style);
            sl.SetCellValue("D4", "I'm a header");

            sl.SaveAs("NamedCellStyles.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
