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

public partial class Workbooks_Formatting_PatternSetting : System.Web.UI.Page
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

        //Get the cells collection
        Cells cells = workbook.Worksheets[0].Cells;

        Aspose.Cells.Style style;

        //Get Style
        style = cells["B1"].GetStyle();

        //Specify the fill color of the cell
        style.ForegroundColor = Color.Red;
        style.Pattern = BackgroundType.Solid;
       
        //Set Style
        cells["B1"].SetStyle(style);

        //Get Style
        style = cells["B2"].GetStyle();

        //Set the background, foreground colors of the cell
        style.ForegroundColor = Color.Yellow;
        style.BackgroundColor = Color.Blue;

        //Set Style Pattern
        style.Pattern = BackgroundType.DiagonalCrosshatch;

        //Set Style
        cells["B2"].SetStyle(style);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "PatternSetting.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "PatternSetting.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }
}
