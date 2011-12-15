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

public partial class Workbooks_DrawingObjects_AddingPictures : System.Web.UI.Page
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
        Worksheet sheet = workbook.Worksheets[0];

        //Insert a picture into a cell
        string ImageUrl = System.Web.HttpContext.Current.Server.MapPath("~/Image/school.JPG");
        int pictureIndex = sheet.Pictures.Add(1, 1, ImageUrl);
        Aspose.Cells.Drawing.Picture picture = sheet.Pictures[pictureIndex];

        //Insert a picture into a cell using a stream
        FileStream fs = File.OpenRead(ImageUrl);
        //Create Byte Type array 
        byte[] data = new Byte[fs.Length];
        //Read Data from stream into array
        fs.Read(data, 0, data.Length);
        //Close Stream
        fs.Close();

        //Crearte Memory Stream Object
        MemoryStream stream = new MemoryStream();
        //Write data in memory
        stream.Write(data, 0, data.Length);

        //Create Image Object and load from stream
        System.Drawing.Image infoImage = System.Drawing.Image.FromStream(stream);
        //Insert a picture into a cell using a stream
        sheet.Pictures.Add(12, 1, stream, 100, 100);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "AddingPictures.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "AddingPictures.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();  
    }
}
