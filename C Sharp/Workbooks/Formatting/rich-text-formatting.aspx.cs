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

public partial class Workbooks_Formatting_RichTextFormatting : System.Web.UI.Page
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
         //Instantiating an Workbook object
        Workbook workbook = new Workbook();

        //Obtaining the reference of the newly added worksheet by passing its sheet index
        Worksheet worksheet = workbook.Worksheets[0];

        //Accessing the "A1" cell from the worksheet
        Aspose.Cells.Cell cell = worksheet.Cells["A1"];

        //Adding some value to the "A1" cell
        cell.PutValue("Rich Text Formatting Demo");

        cell.Characters(3, 15).Font.IsItalic = true;

        cell.Characters(5, 4).Font.Name = "Algerian";

        cell.Characters(21, 4).Font.Color = System.Drawing.Color.Red;

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "RichTextFormatting.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "RichTextFormatting.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }

}



