using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

namespace Aspose.Cells.Demos.Northwind
{
    public partial class SalesByYearSubreportForm : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnExecute_Click(object sender, EventArgs e)
        {
            //Define a workbook to store null initially
            Workbook workbook = null;

            string path = MapPath(".");
            path = path.Substring(0, path.LastIndexOf("\\"));
            
            SalesByYearSubreport salesByYearSubreport = new SalesByYearSubreport(path);
            
            //Create workbook based on the custom method of a class
            workbook = salesByYearSubreport.CreateSalesByYearSubreport();
            
            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "SalesByYearSubreport.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "SalesByYearSubreport.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End(); 

        }
    }
}
