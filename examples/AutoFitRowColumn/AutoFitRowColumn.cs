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

            SLStyle style;

            sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Numbers");

            // The default format code is "General".
            sl.SetCellValue(2, 1, 12345.678909);

            style = sl.CreateStyle();

            style.FormatCode = "#,##0.00";
            sl.SetCellValue(2, 2, 12345.678909);
            sl.SetCellStyle(2, 2, style);

            style.FormatCode = "0.00";
            sl.SetCellValue(2, 3, 5.6789);
            sl.SetCellStyle(2, 3, style);

            style.FormatCode = "$#,##0.00_);[Red]($#,##0.00)";
            sl.SetCellValue(2, 4, -123456789.5678);
            sl.SetCellStyle(2, 4, style);

            style.FormatCode = "_($* #,##0.00_);_($* (#,##0.00);_($* \" - \"??_);_(@_)";
            sl.SetCellValue(2, 5, -123456789.5678);
            sl.SetCellStyle(2, 5, style);
            sl.SetCellValue(3, 5, 123456789.5678);
            sl.SetCellStyle(3, 5, style);

            style.FormatCode = "0.00%";
            sl.SetCellValue(2, 6, 5.6789);
            sl.SetCellStyle(2, 6, style);

            style.FormatCode = "# ?/?";
            sl.SetCellValue(2, 7, 5.6789);
            sl.SetCellStyle(2, 7, style);

            style.FormatCode = "0.000E+00";
            sl.SetCellValue(2, 8, 12345.678909);
            sl.SetCellStyle(2, 8, style);

            sl.SetCellValue(2, 9, true);
            sl.SetCellValue(2, 10, false);

            sl.AutoFitColumn(1, 10);

            sl.AddWorksheet("Dates");

            style.FormatCode = "dd/mm/yyyy";
            sl.SetCellValue(2, 1, new DateTime(2718, 2, 8));
            sl.SetCellStyle(2, 1, style);

            style.FormatCode = "mmmm dd, yyyy";
            sl.SetCellValue(2, 2, new DateTime(2718, 2, 8));
            sl.SetCellStyle(2, 2, style);

            style.FormatCode = "d mmmmm";
            sl.SetCellValue(2, 3, new DateTime(2718, 2, 8));
            sl.SetCellStyle(2, 3, style);

            style.FormatCode = "mmm-yyyy";
            sl.SetCellValue(2, 4, new DateTime(2718, 2, 8));
            sl.SetCellStyle(2, 4, style);

            style.FormatCode = "dd/mm/yyyy h:mm:ss";
            sl.SetCellValue(2, 5, new DateTime(2718, 2, 8, 15, 34, 59));
            sl.SetCellStyle(2, 5, style);

            style.FormatCode = "dd/mm/yyyy h:mm:ss AM/PM";
            sl.SetCellValue(2, 6, new DateTime(2718, 2, 8, 15, 34, 59));
            sl.SetCellStyle(2, 6, style);

            sl.AutoFitColumn(1, 6);

            sl.AddWorksheet("Typeface");

            style = sl.CreateStyle();
            style.FormatCode = "#,##0.00";
            // 30 degree rotation
            style.Alignment.TextRotation = 30;

            sl.SetCellValue(1, 1, 12345.6789);
            sl.SetCellStyle(1, 1, style);

            style = sl.CreateStyle();
            style.SetFont("Perpetua", 24);
            style.Alignment.TextRotation = 30;

            sl.SetCellValue(2, 2, "This is Perpetua");
            sl.SetCellStyle(2, 2, style);

            SLFont ft;
            SLRstType rst;

            style = sl.CreateStyle();
            style.Alignment.TextRotation = 30;

            rst = sl.CreateRstType();
            ft = sl.CreateFont();
            ft.SetFont("Impact", 36);
            rst.AppendText("First Text ", ft);
            ft = sl.CreateFont();
            ft.SetFont("Harrington", 48);
            ft.SetFontThemeColor(SLThemeColorIndexValues.Accent4Color);
            rst.AppendText("Second Text", ft);

            sl.SetCellValue(3, 3, rst);
            sl.SetCellStyle(3, 3, style);

            rst = sl.CreateRstType();
            ft = sl.CreateFont();
            ft.SetFont("Palatino Linotype", 18);
            ft.Underline = UnderlineValues.Single;
            rst.AppendText("First Text ", ft);
            ft = sl.CreateFont();
            ft.SetFont("Consolas", 36);
            ft.Bold = true;
            ft.SetFontThemeColor(SLThemeColorIndexValues.Accent5Color);
            rst.AppendText("Second Text", ft);
            ft = sl.CreateFont();
            ft.SetFont("Rockwell", 16);
            ft.Italic = true;
            ft.Strike = true;
            rst.AppendText("Third Text", ft);

            sl.SetCellValue(4, 4, rst);
            sl.SetCellStyle(4, 4, style);

            sl.AutoFitColumn(1, 4);
            sl.AutoFitRow(1, 4);

            sl.SaveAs("AutoFitRowColumn.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
