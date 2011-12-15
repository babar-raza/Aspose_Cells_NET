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

public partial class Workbooks_Formatting_NumberFormatting : System.Web.UI.Page
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

        //Open template
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\Workbooks\NumberFormatting.xls";


        //Create a new workbook
        Workbook workbook = new Workbook(path);

        //Get the cells collection in the workbook
        Cells cells = workbook.Worksheets[0].Cells;

        Aspose.Cells.Style style;

        //Set number format with built-in index
        for (int i = 1; i < 37; i++)
        {
            cells[i, 1].PutValue(1234.5);

            int Number = cells[i, 0].IntValue;
            
            //Get Style of Cell
            style = cells[i, 1].GetStyle();

            //Set the display number format
            style.Number = Number;

            //Apply Style
            cells[i, 1].SetStyle(style);
        }

        //Set number format with custom format string
        for (int i = 1; i < 4; i++)
        {
            cells[i, 3].PutValue(1234.5);

            //Get Style of Cell
            style = cells[i, 3].GetStyle();

            //Set the display custom number format
            style.Custom = cells[i, 2].StringValue;

            //Apply Style
            cells[i, 3].SetStyle(style);
        }

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "NumberFormatting.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "NumberFormatting.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }

}



