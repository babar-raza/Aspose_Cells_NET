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
    public partial class SalesByCategorySubreportForm : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void btnExecute_Click(object sender, EventArgs e)
        {
            //Define a workbook to store null as its initial value
            Workbook workbook = null;

            string path = MapPath(".");
            path = path.Substring(0, path.LastIndexOf("\\"));
            
            SalesByCategorySubreport salesByCategorySubreport = new SalesByCategorySubreport(path);
            
            //Create the workbook based on the custom method for a class
            workbook = salesByCategorySubreport.CreateSalesByCategorySubreport();
            
            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "SalesByCategorySubreport.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "SalesByCategorySubreport.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End(); 

        }
    }
}


