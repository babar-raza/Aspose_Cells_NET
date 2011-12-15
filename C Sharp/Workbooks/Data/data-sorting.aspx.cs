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

public partial class Workbooks_Data_DataSorting : System.Web.UI.Page
{
    protected System.Web.UI.WebControls.DropDownList ddlFileVersion;

    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void btnExecute_Click(object sender, EventArgs e)
    {
        //Call Method to create report
        CreateStaticReport();
    }

    public void CreateStaticReport()
    {
        //Open template
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\Workbooks\unsorted.xls";


        //Instantiate a new Workbook object.
        Workbook workbook = new Workbook(path);

        //Get the workbook datasorter object.
        DataSorter sorter = workbook.DataSorter;

        //Set the first order for datasorter object.
        sorter.Order1 = Aspose.Cells.SortOrder.Descending;

        //Define the first key.
        sorter.Key1 = 0;

        //Set the second order for datasorter object.
        sorter.Order2 = Aspose.Cells.SortOrder.Ascending;

        //Define the second key.
        sorter.Key2 = 1;

        //Create a cells area (range).
        CellArea ca = new CellArea();

        //Specify the start row index.
        ca.StartRow = 0;

        //Specify the start column index.
        ca.StartColumn = 0;

        //Specify the last row index.
        ca.EndRow = 13;

        //Specify the last column index.
        ca.EndColumn = 1;

        //Sort data in the specified data range (A1:B14)
        sorter.Sort(workbook.Worksheets[0].Cells, ca);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "DataSorting.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "DataSorting.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      

    }

}



