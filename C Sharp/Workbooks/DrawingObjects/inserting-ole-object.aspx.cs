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
using System.IO;
using Aspose.Cells;

public partial class Workbooks_DrawingObjects_InsertingOleObject : System.Web.UI.Page
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
        //Instantiate a new Workbook.
        Workbook workbook = new Workbook();

        //Get the first worksheet. 
        Worksheet sheet = workbook.Worksheets[0];

        //Define a string variable to store the image path.
        string ImageUrl = System.Web.HttpContext.Current.Server.MapPath("~/Image/school.JPG");

        //Get the picture into the streams.
        FileStream fs = File.OpenRead(ImageUrl);

        //Define a byte array.
        byte[] imageData = new Byte[fs.Length];

        //Obtain the picture into the array of bytes from streams.
        fs.Read(imageData, 0, imageData.Length);

        //Close the stream.
        fs.Close();

        //Get an excel file path in a variable.
        string path = System.Web.HttpContext.Current.Server.MapPath("~/designer/book1.xls");

        //Get the file into the streams.
        fs = File.OpenRead(path);

        //Define an array of bytes. 
        byte[] objectData = new Byte[fs.Length];

        //Store the file from streams.
        fs.Read(objectData, 0, objectData.Length);

        //Close the stream.
        fs.Close();

        //Add an Ole object into the worksheet with the image
        //shown in MS Excel.
        sheet.OleObjects.Add(4, 3, 200, 200, imageData);

        //Set embedded ole object data.     
        sheet.OleObjects[0].ObjectData = objectData;  

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "InsertOleObect.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "InsertOleObect.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();  
    }
}
