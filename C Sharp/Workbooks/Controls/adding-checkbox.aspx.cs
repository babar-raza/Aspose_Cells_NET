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

public partial class Workbooks_Controls_AddCheckbox : System.Web.UI.Page
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

    protected void CreateStaticReport()
    {
        //Instantiate a new Workbook.
        Workbook workbook = new Workbook();

        //Add a checkbox to the first worksheet in the workbook.
        int index = workbook.Worksheets[0].CheckBoxes.Add(5, 5, 20, 120);

        //Get the checkbox object.
        Aspose.Cells.Drawing.CheckBox checkbox = workbook.Worksheets[0].CheckBoxes[index];

        //Set its text string.
        checkbox.Text = "Click it!";

        //Put a value into B1 cell.
        workbook.Worksheets[0].Cells["B1"].PutValue("LnkCell");

        //Set B1 cell as a linked cell for the checkbox.
        checkbox.LinkedCell = "B1";

        //Check the checkbox by default.
        checkbox.Value = true;

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "CheckBox.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "CheckBox.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }

}



