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

public partial class Workbooks_Data_DataFilter : System.Web.UI.Page
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
        
        //Put some values into cells 
        cells["A1"].PutValue("Fruit");
        cells["B1"].PutValue("Total");
        cells["A2"].PutValue("Apple");
        cells["B2"].PutValue(1000);
        cells["A3"].PutValue("Orange");
        cells["B3"].PutValue(2500);
        cells["A4"].PutValue("Bananas");
        cells["B4"].PutValue(2500);
        cells["A5"].PutValue("Pear");
        cells["B5"].PutValue(1000);
        cells["A6"].PutValue("Grape");
        cells["B6"].PutValue(2000);

        cells["D1"].PutValue("Count:");
        
        //Set a formula to E1 cell
        cells["E1"].Formula = "=SUBTOTAL(2,B1:B6)";

        //Represents the range to which the specified AutoFilter applies
        sheet.AutoFilter.Range = "A1:B6";

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "DataFilteringAndValidation.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "DataFilteringAndValidation.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      

    }

}



