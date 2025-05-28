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

            // group columns F to I inclusive
            sl.GroupColumns("F", "I");

            // group rows 2 to 5
            sl.GroupRows(2, 5);
            // group rows 3 to 5. This makes rows 3 to 5 on the second level of grouping
            // because of the previous grouping.
            sl.GroupRows(3, 5);

            sl.GroupRows(10, 14);
            sl.GroupRows(16, 18);
            // note that this grouping includes the previous 2 groups
            sl.GroupRows(8, 20);

            // the group is rows 10 to 14, but in Excel, notice that the collapse -/+
            // box is on row 15. Hence we do collapsing/expanding on the row just after
            // the group. Similarly for columns.
            sl.CollapseRows(15);
            // this essentially undoes (undooze? unduhs? Someone tell me the correct pronunciation...)
            // the previous collapse command.
            sl.ExpandRows(15);

            // this collapses the group 16 to 18
            sl.CollapseRows(19);

            sl.GroupRows(24, 27);
            // this ungroups the row 25
            sl.UngroupRows(25, 25);

            sl.SaveAs("GroupRowColumn.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}