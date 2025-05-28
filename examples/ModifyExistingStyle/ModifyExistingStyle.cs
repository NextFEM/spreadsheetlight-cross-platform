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

            SLStyle style = sl.CreateStyle();
            style.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.CornflowerBlue, System.Drawing.Color.CornflowerBlue);

            // set the style to C3
            sl.SetCellStyle(3, 3, style);

            // get the style from C3
            SLStyle modstyle = sl.GetCellStyle("C3");
            modstyle.Border.BottomBorder.BorderStyle = BorderStyleValues.Medium;
            modstyle.Border.BottomBorder.SetBorderThemeColor(SLThemeColorIndexValues.Accent4Color);

            sl.SetCellStyle(5, 5, modstyle);

            sl.SaveAs("ModifyExistingStyle.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
