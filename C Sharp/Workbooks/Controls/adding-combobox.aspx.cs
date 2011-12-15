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

public partial class Workbooks_Controls_AddCombobox : System.Web.UI.Page
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
        //Create a new Workbook.        
        Workbook workbook = new Workbook();

        //Get the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        //Get the worksheet cells collection.
        Cells cells = sheet.Cells;

        //Input a value.
        cells["B3"].PutValue("Employee:");
        Aspose.Cells.Style style = cells["B3"].GetStyle();

        //Set it bold.
        style.Font.IsBold = true;
        cells["B3"].SetStyle(style);

        //Input some values that denote the input range for the combo box.
        cells["A2"].PutValue("Emp001");

        cells["A3"].PutValue("Emp002");

        cells["A4"].PutValue("Emp003");

        cells["A5"].PutValue("Emp004");

        cells["A6"].PutValue("Emp005");

        cells["A7"].PutValue("Emp006");

        //Add a new combo box.
        Aspose.Cells.Drawing.ComboBox comboBox = sheet.Shapes.AddComboBox(2, 0, 2, 0, 22, 100);

        //Set the linked cell;
        comboBox.LinkedCell = "A1";

        //Set the input range.
        comboBox.InputRange = "A2:A7";

        //Set no. of list lines displayed in the combo box's list portion.
        comboBox.DropDownLines = 5;

        //Set the combo box with 3-D shading.
        comboBox.Shadow = true;

        //AutoFit Columns
        sheet.AutoFitColumns();

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "ComboBox.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "ComboBox.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();
    }

}



