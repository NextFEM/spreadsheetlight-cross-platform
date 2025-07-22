using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using SpreadsheetLight.Drawing;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            SLDocument sl = new SLDocument();

            sl.SetCellValue(1, 1, "Percent");
            sl.SetCellValue(1, 2, 0.67);

            sl.SetCellValue(3, 2, "Gauge Data");
            sl.SetCellValue(4, 2, "=MIN(B1,100%)/2");
            sl.SetCellValue(5, 2, "=50%-B4");
            sl.SetCellValue(6, 2, 0.5);

            SLStyle style = new SLStyle();
            style.FormatCode = "0.00%";
            sl.SetCellStyle(1, 2, style);
            sl.SetCellStyle(4, 2, style);
            sl.SetCellStyle(5, 2, style);
            sl.SetCellStyle(6, 2, style);

            SLChart chart = sl.CreateChart("A3", "B6");

            SLPieChartOptions pco = chart.CreatePieChartOptions();
            pco.FirstSliceAngle = 270;
            chart.SetChartType(SLPieChartType.Pie, pco);
            chart.SetChartPosition(1, 2.5, 15, 9.5);
            // 13 pt width, 6 pt height
            chart.Format3D.SetBevelTop(DocumentFormat.OpenXml.Drawing.BevelPresetValues.CoolSlant, 13, 6);
            chart.Fill.SetSolidFill(System.Drawing.Color.Beige, 0);
            chart.HideChartLegend();

            SLDataPointOptions dpo;

            dpo = chart.CreateDataPointOptions();
            dpo.Fill.SetRadialGradient(SLGradientPresetValues.Fire, SLGradientDirectionValues.CenterToTopLeftCorner);
            chart.SetDataPointOptions(1, 1, dpo);

            dpo = chart.CreateDataPointOptions();
            dpo.Fill.SetSolidFill(System.Drawing.Color.LightSkyBlue, 0);
            chart.SetDataPointOptions(1, 2, dpo);

            dpo = chart.CreateDataPointOptions();
            // this effectively makes the data point invisible
            dpo.Fill.SetNoFill();
            dpo.Line.SetNoLine();
            chart.SetDataPointOptions(1, 3, dpo);

            SLFont font = sl.CreateFont();
            font.SetFont("Impact", 24);
            SLRstType rst = sl.CreateRstType();
            string sTextLabel = string.Format("{0}%", (sl.GetCellValueAsDouble(1, 2) * 100.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture));
            rst.AppendText(sTextLabel, font);
            SLDataLabelOptions dlo = chart.CreateDataLabelOptions();
            dlo.SetLabelText(rst);
            chart.SetDataLabelOptions(1, 3, dlo);

            sl.InsertChart(chart);

            sl.SaveAs("ChartGauge.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
