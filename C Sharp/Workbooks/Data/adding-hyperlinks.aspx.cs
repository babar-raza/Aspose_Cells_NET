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

public partial class Workbooks_Data_AddingHyperlinks : System.Web.UI.Page
{
    protected System.Web.UI.WebControls.DropDownList ddlFileVersion;

    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void btnExecute_Click(object sender, EventArgs e)
    {
        //Call Method to create report
        CreateStaticReport();
    }

    public void CreateStaticReport()
    {
        //Create a new Workbook.        
        Workbook workbook = new Workbook();

        //Get the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];

        //Get cells from workbook
        Cells cells = worksheet.Cells;

        //Put a value into a cell
        cells["A1"].PutValue("Visit Aspose");

        //Get Style Object 
        Aspose.Cells.Style style = cells["A1"].GetStyle();

        //Set the font color of the cell to Blue
        style.Font.Color = Color.Blue;

        //Set the font of the cell to Single Underline
        style.Font.Underline = FontUnderlineType.Single;

        //Set the style of A1 cell
        cells["A1"].SetStyle(style);

        //Add a hyperlink to Aspose web sit at cell "A1"
        worksheet.Hyperlinks.Add("A1", 1, 1, "http://www.aspose.com");

        //add a hyperlink to another cell at cell "C1"
        worksheet.Hyperlinks.Add("C1", 1, 1, "Sheet1!A10");

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "AddingHyperlinks.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "AddingHyperlinks.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      

    }
}
