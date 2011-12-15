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
using Aspose.Cells.Drawing;
using Aspose.Cells.Tables;


public partial class CreateListObject : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        
    }
    public static void CreateStaticReport()
    {

        //Open template from path
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\Workbooks\ListObject.xls";


        //Instantiate a new Workbook object.
        Workbook workbook = new Workbook(path);

        //Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        //Get the ListObjects in the first sheet
        ListObjectCollection listObjects = sheet.ListObjects;

        //Add a list object for the given data
        listObjects.Add(1, 1, 13, 5, true);

        //Set the totals visible
        listObjects[0].ShowTotals = true;

        //Add the summary function to the last column in the list
        listObjects[0].ListColumns[4].TotalsCalculation = TotalsCalculation.Sum;

        //Save the excel file
        workbook.Save(HttpContext.Current.Response, "ListObject.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        
        // End response to avoid unneeded html after xls
        HttpContext.Current.Response.End();
       

    }

    protected void btnExecute_Click(object sender, EventArgs e)
    {
        CreateStaticReport();
    }
}
