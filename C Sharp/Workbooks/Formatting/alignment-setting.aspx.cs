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

public partial class Workbooks_Formatting_AlignmentSetting : System.Web.UI.Page
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
        //Open template from path
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\Workbooks\AlignmentSetting.xls";


        //Create a new workbook
        Workbook workbook = new Workbook(path);

        //Get the cells collection in the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        //Get Style Object 
        Aspose.Cells.Style style = cells["A1"].GetStyle();

        //Set text alignment type
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;

        //Set A1 style
        cells["A1"].SetStyle(style);

        //Get Style Object 
        style = cells["A2"].GetStyle();

        //Set text rotation angel
        style.RotationAngle = 45;

        //Set A2 style
        cells["A2"].SetStyle(style);

        //Get Style Object 
        style = cells["C3"].GetStyle();

        //Set shrinktofit on
        style.ShrinkToFit = true;

        //Set A3 style
        cells["C3"].SetStyle(style);

        //Get Style Object 
        style = cells["A4"].GetStyle();

        //Set the indentlevel
        style.IndentLevel = 5;

        //Set A4 style
        cells["A4"].SetStyle(style);

        //Get Style Object 
        style = cells["A5"].GetStyle();

        //Wrapping Text
        style.IsTextWrapped = true;

        //Set A5 style
        cells["A5"].SetStyle(style);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "AlignmentSetting.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "AlignmentSetting.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }
}



