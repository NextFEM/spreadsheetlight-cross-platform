using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace WebApp
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            // "Response" refers to Page.Response, an HttpResponse class
            SLDocument sl = new SLDocument();
            sl.SetCellValue(2, 2, "I'm going on the Internet super highway!");
            sl.ApplyNamedCellStyle(2, 2, SLNamedCellStyleValues.Good);
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=WebStreamDownload.xlsx");
            sl.SaveAs(Response.OutputStream);
            Response.End();
        }
    }
}
