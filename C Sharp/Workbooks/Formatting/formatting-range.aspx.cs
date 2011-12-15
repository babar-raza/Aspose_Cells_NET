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

public partial class Workbooks_Formatting_FormattingRange : System.Web.UI.Page
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
        //Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        //Get its cells collection
        Cells cells = sheet.Cells;

        //Create a named range
        Range range = sheet.Cells.CreateRange("B1", "E5");
        //Set the name of the named range
        range.Name = "Range1";
        //Create a new style adding to the workbook styles collection
        Aspose.Cells.Style style = workbook.Styles[workbook.Styles.Add()];
        //Specify the style's fill color
        style.ForegroundColor = System.Drawing.Color.Blue;
        style.Pattern = BackgroundType.Solid;

        //Create a styleflag object
        StyleFlag styleFlag = new StyleFlag();
        //Specify all attributes
        styleFlag.All = true;
        //Apply the style to the range
        range.ApplyStyle(style, styleFlag);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "FormattingRange.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "FormattingRange.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      

    }

}



