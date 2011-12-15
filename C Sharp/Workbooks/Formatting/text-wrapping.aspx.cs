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

public partial class Workbooks_Formatting_TextWrapping : System.Web.UI.Page
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
        //Create Workbook Object
        Workbook wb = new Workbook();

        //Open first Worksheet in the workbook
        Worksheet ws = wb.Worksheets[0];

        //Get Worksheet Cells Collection
        Aspose.Cells.Cells cell = ws.Cells;

        //Increase the width of First Column Width
        cell.SetColumnWidth(0, 35);

        //Increase the height of first row
        cell.SetRowHeight(0, 36);

        //Add Text to the Firts Cell
        cell[0, 0].PutValue("This is the example of text wrap functionality using Aspose.Cells.");

        //Get Style
        Aspose.Cells.Style style = cell[0, 0].GetStyle();

        //Make Cell's Text wrap
        style.IsTextWrapped = true;

        //Set Style
        cell[0, 0].SetStyle(style);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            wb.Save(HttpContext.Current.Response, "TextWrapping.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            wb.Save(HttpContext.Current.Response, "TextWrapping.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }

}



