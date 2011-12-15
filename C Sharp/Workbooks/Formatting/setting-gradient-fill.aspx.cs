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
using System.Drawing;
using Aspose.Cells;

public partial class Workbooks_Formatting_GradientFill : System.Web.UI.Page
{
    protected System.Web.UI.WebControls.DropDownList ddlFileVersion;

    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void btnExecute_Click(object sender, EventArgs e)
    {
        CreateStaticReport();
    }

    public void CreateStaticReport()
    {
        //Create a new workbook
        Workbook workbook = new Workbook();

        //Get first worksheet in the workbook
        Worksheet sheet = workbook.Worksheets[0];

        //Get cell A1 from worksheet's cell collection
        Aspose.Cells.Cell cell = sheet.Cells["A1"];

        //Get style of the cell
        Aspose.Cells.Style style = cell.GetStyle();

        //Set Two Color Gradient
        style.SetTwoColorGradient(Color.Red, Color.Green, Aspose.Cells.Drawing.GradientStyleType.Horizontal, 1);

        //Apply cell style
        cell.SetStyle(style);

        //Set row height and column width
        sheet.Cells.SetColumnWidth(0, 50);
        sheet.Cells.SetRowHeight(0, 50);

        
        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "GradientFill.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "GradientFill.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }
}



