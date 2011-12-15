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
using System.IO;

public partial class Workbooks_DrawingObjects_AddingImageHyperlink : System.Web.UI.Page
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

        //Create worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        //Insert a picture into a cell
        string ImageUrl = System.Web.HttpContext.Current.Server.MapPath("~/Image/school.jpg");
      
        //Insert a string value to a cell
        worksheet.Cells["C2"].PutValue("Image Hyperlink");

        //Set the 4th row height
        worksheet.Cells.SetRowHeight(3, 100);

        //Set the C column width
        worksheet.Cells.SetColumnWidth(2, 21);

        //Add a picture to the C4 cell
        int index = worksheet.Pictures.Add(3, 2, 4, 3, ImageUrl);

        //Get the picture object
        Aspose.Cells.Drawing.Picture pic = worksheet.Pictures[index];

        //Set the placement type
        pic.Placement = Aspose.Cells.Drawing.PlacementType.FreeFloating;

        //Add an image hyperlink
        Aspose.Cells.Hyperlink hlink = pic.AddHyperlink("http://www.aspose.com/");

        //Specify the screen tip
        hlink.ScreenTip = "Click to go to Aspose site";

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "AddingImageHyperlink.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "AddingImageHyperlink.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();  
    }
}
