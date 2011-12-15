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
    public partial class ProductsByCategoryForm : System.Web.UI.Page
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
            
            ProductsByCategory productsByCategory = new ProductsByCategory(path);
            
            //Create workbook based on a custom method of a class
            workbook = productsByCategory.CreateProductsByCategory();
            
            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "ProductsByCategory.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "ProductsByCategory.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End(); 

        }
    }
}
