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

public partial class Copy_Move : System.Web.UI.Page
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
        path += @"\designer\Workbooks\Copy_Move.xls";

        //Instantiate a new Workbook object.
        Workbook workbook = new Workbook(path);

        //Copy the first sheet contents into the last worksheet in the book
        workbook.Worksheets[2].Copy(workbook.Worksheets["Copy"]);

        //Move the sheet to the last indexed position in the book 
        workbook.Worksheets["Move"].Move(2);

        //Save the excel file
        workbook.Save(HttpContext.Current.Response, "CopyandMove.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        
        // End response to avoid unneeded html after xls
        HttpContext.Current.Response.End();
    }
}
