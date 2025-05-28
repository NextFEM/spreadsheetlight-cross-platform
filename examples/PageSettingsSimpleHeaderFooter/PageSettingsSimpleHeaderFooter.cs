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

            SLPageSettings ps = new SLPageSettings();

            ps.SetLeftHeaderText("I'm left-handed");
            ps.SetRightFooterText("I put the right foot forward");

            // the default is to set on odd-numbered pages, but you can specify otherwise
            // such as on the first page or on even-numbered pages.
            ps.SetCenterHeaderText(SLHeaderFooterTypeValues.First, "I'm the head");
            // you'll have to set this true to see the first page settings
            ps.DifferentFirstPage = true;

            // or this if you want odd-even pages to be seen.
            //ps.DifferentOddEvenPages = true;

            sl.SetPageSettings(ps);

            sl.SaveAs("PageSettingsSimpleHeaderFooter.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
