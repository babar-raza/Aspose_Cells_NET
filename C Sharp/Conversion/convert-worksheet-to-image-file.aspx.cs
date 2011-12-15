using System;
using System.Data;
using System.IO;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Aspose.Cells;
using Aspose.Cells.Rendering;

public partial class Sheet2Image : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {


    }
    protected void btnExecute_Click(object sender, EventArgs e)
    {
        CreateStaticReport();
    }

    public static void CreateStaticReport()
    {
        //Open template
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\MyTestBook1.xls";

        //Instantiate a new Workbook object.
        Workbook book = new Workbook(path);

        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
        
        Worksheet sheet = book.Worksheets[0];
        SheetRender sheetRender = new SheetRender(sheet, imgOptions);
       
        //Create a memory stream object.
        MemoryStream memorystream = new MemoryStream();

        //Convert worksheet to image.
        sheetRender.ToImage(0, memorystream);

        memorystream.Seek(0, SeekOrigin.Begin);  

        //Set Response object to stream the image file.
        byte[] data = memorystream.ToArray();
        HttpContext.Current.Response.Clear();
        HttpContext.Current.Response.ContentType = "image/jpeg";
        HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=SheetImage.jpeg");
        HttpContext.Current.Response.OutputStream.Write(data, 0, data.Length);

        //End response to avoid unneeded html after xls
        HttpContext.Current.Response.End();

    }


}
