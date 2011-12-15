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
using System.Drawing;

public partial class Workbooks_Formatting_BorderSetting : System.Web.UI.Page
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
        //Get the cells collection in the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        //Get Style of B2
        Aspose.Cells.Style style = cells["B2"].GetStyle();

        //Set the cell border color
        style.Borders[BorderType.TopBorder].Color = Color.Blue;
        style.Borders[BorderType.BottomBorder].Color = Color.Blue;
        style.Borders[BorderType.LeftBorder].Color = Color.Blue;
        style.Borders[BorderType.RightBorder].Color = Color.Blue;
        style.Borders[BorderType.DiagonalDown].Color = Color.Blue;
        style.Borders[BorderType.DiagonalUp].Color = Color.Blue;


        //Set the cell border type
        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.DashDot;
        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.DashDot;
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.DashDot;
        style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.DashDot;
        style.Borders[BorderType.DiagonalDown].LineStyle = CellBorderType.DashDot;
        style.Borders[BorderType.DiagonalUp].LineStyle = CellBorderType.DashDot;

        //Setting Border Style for B2
        cells["B2"].SetStyle(style);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "BorderSetting.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "BorderSetting.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }
}



