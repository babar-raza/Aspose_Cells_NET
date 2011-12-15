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

public partial class Workbooks_RowsAndColumns_AdjustingRowsAndColumns : System.Web.UI.Page
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

        Cells cells = workbook.Worksheets[0].Cells;

        //Set the height of all row in the worksheet
        cells.StandardHeight = 20;

        //Set the width of all columns in the worksheet
        cells.StandardWidth = 20;

        //Set the width of the first column 
        cells.SetColumnWidth(0, 12);

        //Set the width of the column 
        cells.SetColumnWidth(1, 40);
        //Set the height of the row 
        cells.SetRowHeight(1, 8);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "RowHeightandColumnWidth.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "RowHeightandColumnWidth.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }

}
