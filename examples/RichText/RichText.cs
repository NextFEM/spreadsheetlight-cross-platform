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

            SLFont redunderline = sl.CreateFont();
            redunderline.SetFont(FontSchemeValues.Minor, 18);
            redunderline.FontColor = System.Drawing.Color.IndianRed;
            redunderline.Underline = UnderlineValues.Double;

            SLFont greenbold = sl.CreateFont();
            greenbold.SetFont("Harrington", 24);
            greenbold.Bold = true;
            greenbold.FontColor = System.Drawing.Color.Green;

            SLRstType rst = sl.CreateRstType();
            rst.AppendText("This is in red and underlined, ", redunderline);
            rst.AppendText("and this is in green and bold.", greenbold);

            sl.SetCellValue(3, 3, rst.ToInlineString());

            rst = sl.CreateRstType();
            rst.AppendText("First! ", redunderline);
            rst.AppendText("Second!", greenbold);

            SLFont blueitalic = sl.CreateFont();
            blueitalic.Italic = true;
            blueitalic.FontColor = System.Drawing.Color.CadetBlue;
            List<SLRun> listruns = rst.GetRuns();
            listruns.Insert(1, new SLRun() { Font = blueitalic, Text = "Imma blue y'all. " });
            // because of insertion, the 2nd item (index 1) is now the 3rd (index 2).
            listruns[2].Text = listruns[2].Text.Replace("Second", "Third");
            rst.ReplaceRuns(listruns);

            sl.SetCellValue(5, 3, rst);

            sl.SaveAs("RichText.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
