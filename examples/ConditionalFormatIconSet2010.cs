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

            Random rand = new Random();
            int i, j;
            for (i = 1; i <= 20; ++i)
            {
                for (j = 1; j <= 10; ++j)
                {
                    sl.SetCellValue(i, j, 200 * rand.NextDouble());
                }
            }

            SLThreeIconSetOptions iso3;
            SLFourIconSetOptions iso4;
            SLFiveIconSetOptions iso5;
            SLConditionalFormatting cf;

            iso3 = new SLThreeIconSetOptions(SLThreeIconSetValues.ThreeStars);
            cf = new SLConditionalFormatting("B3", "D11");
            cf.SetCustomIconSet(iso3);
            sl.AddConditionalFormatting(cf);

            iso4 = new SLFourIconSetOptions(SLFourIconSetValues.FourTrafficLights);
            // manipulate the icons!
            iso4.Icon1 = SLIconValues.NoIcon;
            iso4.Icon3 = SLIconValues.YellowFlag;
            cf = new SLConditionalFormatting("F2", "I10");
            cf.SetCustomIconSet(iso4);
            sl.AddConditionalFormatting(cf);

            iso5 = new SLFiveIconSetOptions(SLFiveIconSetValues.FiveBoxes);
            iso5.Icon3 = SLIconValues.GreenUpTriangle;
            // you can also change the type, but it's too much effort to come up with a valid example...
            // For example, you can do something like:
            //iso5.Type2 = SLConditionalFormatRangeValues.Number;
            // The default is Percent.
            iso5.Value2 = "15";
            cf = new SLConditionalFormatting("E12", "J18");
            cf.SetCustomIconSet(iso5);
            sl.AddConditionalFormatting(cf);

            sl.SaveAs("ConditionalFormatIconSet2010.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
