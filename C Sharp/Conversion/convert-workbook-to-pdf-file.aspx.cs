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

public partial class Xls2Pdf : System.Web.UI.Page
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

        //Save the workbook as a PDF File
        book.Save(HttpContext.Current.Response, "Xls2Pdf.pdf", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Pdf));

        //End response to avoid unneeded html after xls
        HttpContext.Current.Response.End();

    }


}
