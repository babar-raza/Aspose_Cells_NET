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

public partial class Workbooks_Formatting_SuperscriptSubscript : System.Web.UI.Page
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
        //Instantiating a Workbook object
        Workbook workbook = new Workbook();
   
        //Obtaining the reference to the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        //Accessing the "A1" cell from the worksheet
        Cell cell = worksheet.Cells["A1"];

        //Get Style
        Aspose.Cells.Style style = cell.GetStyle();
        
        //Adding some value to the "A1" cell
        cell.PutValue("Hello");

        //Setting the font Superscript
        style.Font.IsSuperscript = true;

        //Set Style
        cell.SetStyle(style);

        //Get Cell
        cell = worksheet.Cells["A2"];

        //Get Style
        style = cell.GetStyle();

        //Adding some value to the "A2" cell
        cell.PutValue("Aspose");

        //Setting the font Superscript
        style.Font.IsSubscript = true;

        //Set Style
        cell.SetStyle(style);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "SuperscriptSubscript.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "SuperscriptSubscript.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }

}



