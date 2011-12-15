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

public partial class Workbooks_DrawingObjects_AddingComments : System.Web.UI.Page
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
        //Create Workbook
        Workbook workbook = new Workbook();

        //Create Worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        //Create Cells
        Cells cells = worksheet.Cells;
        
        //Put a value into a cell
        cells["B1"].PutValue("Hello");

        //Add comment to cell B1
        int commentIndex = worksheet.Comments.Add(0, 1);
        
        //Access the newly added comment
        Comment comment = worksheet.Comments[commentIndex];
        
        //Set the comment note
        comment.Note = "Aspose.Cells";

        //Set the font of a comment
        comment.Font.Size = 12;
        comment.Font.IsBold = true;
        comment.HeightCM = 5;
        comment.WidthCM = 5;

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "AddingComments.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "AddingComments.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();  
    }
}
