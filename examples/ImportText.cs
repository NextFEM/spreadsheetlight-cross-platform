using System;
using System.Collections.Generic;
using System.Globalization;
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
            using (SLDocument sl = new SLDocument())
            {
                sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Delimited");

                SLTextImportOptions tio = new SLTextImportOptions();

                // By default, only the tab delimiter is used. If you need comma
                // as a delimiter, set like so:
                //tio.UseCommaDelimiter = true;
                // Speaking of which...

                // By default, the data source is assumed to be character delimited.
                // If the data source is fixed width, do this:
                //tio.DataFieldType = SLTextImportDataFieldTypeValues.FixedWidth;

                // Use the German culture. Because why not?
                // Ok because I need to show the culture-specific parsing and stuff...
                tio.Culture = new CultureInfo("de-DE");

                tio.SetColumnFormat(1, SLTextImportColumnFormatValues.Text);
                // don't set column format type for column 2, so it's General
                // Set the column format type as Text if you want to force it to be text.
                // Basically, whatever you do in Excel or LibreOffice Calc, you do the
                // same here.

                // You can set as General specifically too.
                tio.SetColumnFormat(3, SLTextImportColumnFormatValues.General);
                // There's a Skip option in case you don't want specific columns.
                // Here's a shortcut function:
                //tio.SkipColumn(3);

                // The 4th data column has random data, in text, in number and in date.
                // We leave as General type so you can see what happens.
                // However, no style formatting is done for date data, so you still have
                // to explicitly set a date format using SLStyle.

                tio.SetColumnFormat(5, SLTextImportColumnFormatValues.DateDMY);
                
                // The 6th column has the exact same date text data as the 5th column.
                // This is to show how the parsing is done during the import for a
                // different date part order.
                tio.SetColumnFormat(6, SLTextImportColumnFormatValues.DateMDY);

                // If you set a column as a date format type, SpreadsheetLight will do
                // its best to follow the order of the date parts.
                // For example, DateDMY means "5/6/2013", "05 June 2013", "5 Jun 2013"
                // "5.6.2013" all mean the 5th of June in the year 2013.
                tio.SetColumnFormat(7, SLTextImportColumnFormatValues.DateDMY);

                // Use this in case you have dates in weird text formats...
                tio.AddCustomDateFormat("dd || MMMM -=- yyyy");
                // You still need to state that you want this particular column as a date though...
                tio.SetColumnFormat(8, SLTextImportColumnFormatValues.DateDMY);

                sl.ImportText("ImportTextDelimited.txt", "B3", tio);

                SLStyle style = sl.CreateStyle();
                style.FormatCode = "dd MMM yyyy";

                sl.SetColumnStyle(6, 9, style);

                sl.AutoFitColumn(2, 9);

                // Ok fine, let's do one example with fixed width too.
                sl.AddWorksheet("FixedWidth");

                tio = new SLTextImportOptions(SLTextImportDataFieldTypeValues.FixedWidth);
                tio.Culture = new CultureInfo("de-DE");

                // most of the stuff is the same as before, so I'll just show the differences.

                // this sets the 1st column to be 4 characters wide
                tio.SetFixedWidth(1, 4);
                // this sets the 2nd column to be 5 characters wide
                tio.SetFixedWidth(2, 5);
                // this sets the 3rd column to be 6 characters wide
                tio.SetFixedWidth(3, 6);
                // this sets... oh you know the drill...
                tio.SetFixedWidth(4, 10);

                // By default, spaces are preserved. If after fixed-width-separation,
                // the data is say "a   " (note the spaces), they are kept.
                tio.PreserveSpace = false;
                // Set the above property to false and SpreadsheetLight will .Trim()
                // the fudgecake out of every column data.

                // Any column that doesn't have an explicit width assigned is
                // 8 characters wide by default. But you can change that with
                //tio.DefaultFixedWidth = 32;

                tio.SetColumnFormat(4, SLTextImportColumnFormatValues.DateDMY);

                sl.ImportText("ImportTextFixedWidth.txt", "B3", tio);

                // reuse the date format SLStyle from before
                sl.SetColumnStyle(5, style);

                sl.AutoFitColumn(2, 5);

                sl.SaveAs("ImportText.xlsx");
            }

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
