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

public partial class Sheet2ImageWithOptions : System.Web.UI.Page
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

        //Get the first worksheet
        Worksheet sheet = book.Worksheets[0];

        //Apply different Image and Print options 
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.HorizontalResolution = 300;
        options.VerticalResolution = 300;
        options.TiffCompression = TiffCompression.CompressionCCITT4;
        options.IsCellAutoFit = false;
        options.ImageFormat = System.Drawing.Imaging.ImageFormat.Tiff;
        options.PrintingPage = PrintingPageType.Default;

        //Create a memory stream object.
        MemoryStream memorystream = new MemoryStream();
       
        SheetRender sheetRender = new SheetRender(sheet, options);

        //Convert worksheet to image.
        sheetRender.ToTiff(memorystream);

        memorystream.Seek(0, SeekOrigin.Begin);  
        
        //Set Response object to stream the image file.
        byte[] data = memorystream.ToArray();
        HttpContext.Current.Response.Clear();
        HttpContext.Current.Response.ContentType = "image/tiff";
        HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=SheetImage.tiff");
        HttpContext.Current.Response.OutputStream.Write(data, 0, data.Length);

        //End response to avoid unneeded html after xls
        HttpContext.Current.Response.End();

    }


}
