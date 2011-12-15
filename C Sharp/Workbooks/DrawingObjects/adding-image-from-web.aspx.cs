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

public partial class Workbooks_DrawingObjects_AddingImageFromWeb : System.Web.UI.Page
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
        //Define memory stream object
        System.IO.MemoryStream objImage;

        //Define web client object
        System.Net.WebClient objwebClient;

        //Define a string which will hold the web image url
        string sURL = "http://www.xlsoft.com/jp/products/aspose/images/Aspose_Cells-Product-Box.jpg";
              
        //Instantiate the web client object
        objwebClient = new System.Net.WebClient();

        //Now, extract data into memory stream downloading the image data into the array of bytes
        objImage = new System.IO.MemoryStream(objwebClient.DownloadData(sURL));

        //Create a new workbook
        Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

        //Get the first worksheet in the book
        Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

        //Get the first worksheet pictures collection
        Aspose.Cells.Drawing.PictureCollection pictures = sheet.Pictures;

        //Insert the picture from the stream to B2 cell
        pictures.Add(1, 1, objImage);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "AddingImageFromWeb.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "AddingImageFromWeb.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();  
    }
}
