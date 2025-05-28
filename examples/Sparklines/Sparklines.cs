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
            for (int i = 2; i <= 6; ++i)
            {
                for (int j = 2; j <= 8; ++j)
                {
                    sl.SetCellValue(i, j, 200.0 * rand.NextDouble() - 100.0);
                }
            }

            SLStyle numstyle = sl.CreateStyle();
            numstyle.FormatCode = "0.00";
            sl.SetCellStyle(2, 2, 6, 8, numstyle);

            SLSparklineGroup spkgrp;

            // SetLocation() needs a 1-dimensional vector, meaning it's either a single row
            // or single column. And the length of that vector needs to be equal to
            // either the number of rows or number of columns in the data source range.

            // Excel has a user interface to pop up error dialog boxes. We don't have this
            // luxury, so just be careful...

            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("I2", "I6");
            sl.InsertSparklineGroup(spkgrp);

            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("K2", "O2");
            sl.InsertSparklineGroup(spkgrp);

            // notice that the length of the location range determines how the data is arranged,
            // whether the data series is in rows or is in columns.

            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("B7", "H7");
            sl.InsertSparklineGroup(spkgrp);

            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("B9", "B15");
            sl.InsertSparklineGroup(spkgrp);

            // set sparkline style
            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("E10", "I10");
            spkgrp.SetSparklineStyle(SLSparklineStyle.Accent5Darker25Percent);
            sl.InsertSparklineGroup(spkgrp);

            // set sparkline colours
            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("E11", "I11");
            spkgrp.SeriesColor.Color = System.Drawing.Color.DarkKhaki;
            spkgrp.NegativeColor.Color = System.Drawing.Color.Red;
            spkgrp.MarkersColor.Color = System.Drawing.Color.Yellow;
            spkgrp.HighMarkerColor.Color = System.Drawing.Color.Green;
            spkgrp.LowMarkerColor.Color = System.Drawing.Color.Blue;
            spkgrp.FirstMarkerColor.Color = System.Drawing.Color.Indigo;
            // use a theme colour just because I can. :)
            spkgrp.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
            sl.InsertSparklineGroup(spkgrp);
            // of course, those colours won't appear because the related points aren't shown.
            // Speaking of which...

            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("E12", "I12");
            spkgrp.ShowNegativePoints = true;
            spkgrp.ShowMarkers = true;
            spkgrp.ShowHighPoint = true;
            spkgrp.ShowLowPoint = true;
            spkgrp.ShowFirstPoint = true;
            spkgrp.ShowLastPoint = true;
            // show hidden data, but we don't have hidden cells so whatever...
            spkgrp.ShowHiddenData = true;
            spkgrp.ShowEmptyCellsAs = DocumentFormat.OpenXml.Office2010.Excel.DisplayBlanksAsValues.Zero;
            sl.InsertSparklineGroup(spkgrp);

            sl.SetCellValue(2, 1, new DateTime(3456, 8, 15));
            sl.SetCellValue(3, 1, new DateTime(3456, 8, 16));
            sl.SetCellValue(4, 1, new DateTime(3456, 8, 17));
            sl.SetCellValue(5, 1, new DateTime(3456, 8, 18));
            sl.SetCellValue(6, 1, new DateTime(3456, 8, 19));
            SLStyle datestyle = sl.CreateStyle();
            datestyle.FormatCode = "d-mmm";
            sl.SetCellStyle("A2", "A6", datestyle);

            // set horizontal axis as a date axis. We'll need a data source range with the dates
            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("E13", "I13");
            spkgrp.SetDateAxis("A2", "A6");
            sl.InsertSparklineGroup(spkgrp);

            // the horizontal axis only shows if the zero line is crossed.
            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("E14", "I14");
            spkgrp.ShowAxis = true;
            sl.InsertSparklineGroup(spkgrp);

            // set right-to-left
            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("E15", "I15");
            spkgrp.RightToLeft = true;
            sl.InsertSparklineGroup(spkgrp);

            // by default the minimum and maximum values are automatic. Or
            //spkgrp.SetAutomaticMinimumValue();
            //spkgrp.SetAutomaticMaximumValue();
            // Speaking of which...

            // set the same minimum value for the whole group
            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("E16", "I16");
            spkgrp.SetSameMinimumValue();
            sl.InsertSparklineGroup(spkgrp);

            // set the same maximum value for the whole group
            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("E17", "I17");
            spkgrp.SetSameMaximumValue();
            sl.InsertSparklineGroup(spkgrp);

            // set a custom minimum value for the whole group
            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("E18", "I18");
            spkgrp.SetCustomMinimumValue(-180);
            sl.InsertSparklineGroup(spkgrp);

            // set a custom maximum value for the whole group
            spkgrp = sl.CreateSparklineGroup("B2", "H6");
            // the default sparkline type is "Line"
            spkgrp.SetLocation("E19", "I19");
            spkgrp.SetCustomMaximumValue(180);
            sl.InsertSparklineGroup(spkgrp);

            sl.SaveAs("Sparklines.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
