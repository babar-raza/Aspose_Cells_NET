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

public partial class Workbook2Image : System.Web.UI.Page
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
        path += @"\designer\FinancialPlan.xls";
    

        Workbook workbook = new Workbook(path);

      

        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

        imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Tiff;

        imgOptions.HorizontalResolution = 100;

        imgOptions.VerticalResolution = 100;

        imgOptions.OnePagePerSheet = true;

        WorkbookRender bookRender = new WorkbookRender(workbook, imgOptions);

        //Create a memory stream object.
        MemoryStream memorystream = new MemoryStream();

        bookRender.ToImage(memorystream);

        memorystream.Seek(0, SeekOrigin.Begin);

        //Set Response object to stream the image file.
        byte[] data = memorystream.ToArray();
        HttpContext.Current.Response.Clear();
        HttpContext.Current.Response.ContentType = "image/tiff";
        HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=WorkbookImage.tiff");
        HttpContext.Current.Response.OutputStream.Write(data, 0, data.Length);

        //End response to avoid unneeded html after xls
        HttpContext.Current.Response.End();
    }


}
