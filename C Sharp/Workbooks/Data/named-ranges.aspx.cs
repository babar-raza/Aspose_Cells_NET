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

public partial class Workbooks_Data_NamedRanges : System.Web.UI.Page
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
        //Create a new workbook
        Workbook workbook = new Workbook();
        
        //Get the first worksheet in the workbook
        Worksheet sheet = workbook.Worksheets[0];

        //Get the cells collection in the sheet
        Cells cells = sheet.Cells;

        //Create a named range
        Range range = cells.CreateRange("B1", "E5");
        
        //Set the name of the named range
        range.Name = "TestRange";

        //Accessing a specific Named Range
        Range myRange = workbook.Worksheets.GetRangeByName("TestRange");
        
        //Get the first cell in the range
        Aspose.Cells.Cell cell = myRange[0, 0];
        
        //Put string value to it
        cell.PutValue("Top left of TestRange");
        
        //Get Style Object 
        Aspose.Cells.Style style = cell.GetStyle();
        
        //Set the fill color of the cell
        style.ForegroundColor = System.Drawing.Color.Blue;
        style.Pattern = BackgroundType.Solid;
        cell.SetStyle(style);

        //Get the last cell in the range
        cell = myRange[myRange.RowCount - 1, myRange.ColumnCount - 1];
        
        //Put a string value to it
        cell.PutValue("Bottom right of TestRange");

        //Set the fill color of the cell
        style.ForegroundColor = System.Drawing.Color.Blue;
        style.Pattern = BackgroundType.Solid;
        cell.SetStyle(style);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "NamedRanges.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "NamedRanges.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      

    }
}

