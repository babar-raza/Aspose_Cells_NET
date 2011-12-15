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

public partial class Workbooks_DrawingObjects_OtherDrawingObjects : System.Web.UI.Page
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
        Worksheet sheet = workbook.Worksheets[0];

        //Add Textbox object in collection
        int textboxIndex = sheet.TextBoxes.Add(1, 1, 40, 40);

        //Get newly added Textbox from collection
        Aspose.Cells.Drawing.TextBox textbox = sheet.TextBoxes[textboxIndex];
        
        //Set TextBox Text
        textbox.Text = "Sample Text Box";

        //Set Textbox dimensions
        textbox.Height = 80;
        textbox.Width = 80;

        //Get path of Image in Variable
        string imageUrl = System.Web.HttpContext.Current.Server.MapPath("~/Image/school.jpg");
        
        //Create File Stream to read image Data
        FileStream fs = File.OpenRead(imageUrl);

        //Initialize Byte Array to store Image Data
        byte[] imageData = new Byte[fs.Length];

        //Read File Stream Data into Array
        fs.Read(imageData, 0, imageData.Length);

        //Cloese File Stream
        fs.Close();

        //Open template
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\book1.xls";

        //Read Template file through Stream
        fs = File.OpenRead(path);

        //Create Byte array to store Template file data
        byte[] objectData = new Byte[fs.Length];

        //Start read Data
        fs.Read(objectData, 0, objectData.Length);

        //Close File Stream
        fs.Close();

        //Add Image as Ole Objects to Worksheet OleObjects Collection
        sheet.OleObjects.Add(3, 3, 150, 150, imageData);

        // embedded ole object data 
        sheet.OleObjects[0].ObjectData = objectData;  

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "OtherDrawingObjects.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "OtherDrawingObjects.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();  

    }
}
