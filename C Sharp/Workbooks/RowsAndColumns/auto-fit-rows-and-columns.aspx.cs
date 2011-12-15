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
using Aspose.Cells;

public partial class Workbooks_RowsAndColumns_AutoFitRowsAndColumns : System.Web.UI.Page
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
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        Cells cells = sheet.Cells;

        cells["B1"].PutValue("Aspose.Cells");
        //Get Style Object 
        Aspose.Cells.Style style = cells["B1"].GetStyle();

        style.RotationAngle = 45;
        style.Font.IsBold = true;
        cells["B1"].SetStyle(style);

        //Auto row fit
        sheet.AutoFitRow(0);
        //Auto column fit
        sheet.AutoFitColumn(1);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "AutoFitRowsAndColumns.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "AutoFitRowsAndColumns.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }
}
