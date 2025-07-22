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

            sl.InsertHyperlink("B3", SLHyperlinkTypeValues.Url, "http://spreadsheetlight.com/");
            
            // can also be a defined name
            sl.InsertHyperlink("B5", SLHyperlinkTypeValues.InternalDocumentLink, "A2");

            // don't include the "mailto:" part
            sl.InsertHyperlink("B7", SLHyperlinkTypeValues.EmailAddress, "immadisturbpresident@abcdef.com");

            // I can't show you a link to an internal document, but here's an example:
            //sl.InsertHyperlink("B9", SLHyperlinkTypeValues.FilePath, "C:\\mahsekretfahl.txt");

            sl.SaveAs("Hyperlinks.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
