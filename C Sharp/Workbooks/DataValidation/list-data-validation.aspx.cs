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
using System.Drawing;

public partial class ListDataValidation : System.Web.UI.Page
{
    protected System.Web.UI.WebControls.DropDownList ddlFileVersion;

    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void btnExecute_Click(object sender, EventArgs e)
    {
        CreateStaticReport();
    }

    public void CreateStaticReport()
    {
        // Create a workbook object.
        Workbook workbook = new Workbook();

        // Get the first worksheet.
        Worksheet worksheet1 = workbook.Worksheets[0];

        // Add a new worksheet and access it.
        int i = workbook.Worksheets.Add();

        Worksheet worksheet2 = workbook.Worksheets[i];

        // Create a range in the second worksheet.
        Range range = worksheet2.Cells.CreateRange("E1", "E4");

        // Name the range.
        range.Name = "MyRange";

        // Fill different cells with data in the range.
        range[0, 0].PutValue("Blue");
        range[1, 0].PutValue("Red");
        range[2, 0].PutValue("Green");
        range[3, 0].PutValue("Yellow");

        // Get the validations collection.
        ValidationCollection validations = worksheet1.Validations;

        // Create a new validation to the validations list.
        Validation validation = validations[validations.Add()];

        // Set the validation type.
        validation.Type = Aspose.Cells.ValidationType.List;

        // Set the operator.
        validation.Operator = OperatorType.None;

        // Set the in cell drop down.
        validation.InCellDropDown = true;

        // Set the formula1.
        validation.Formula1 = "=MyRange";

        // Enable it to show error.
        validation.ShowError = true;

        // Set the alert type severity level.
        validation.AlertStyle = ValidationAlertType.Stop;

        // Set the error title.
        validation.ErrorTitle = "Error";

        // Set the error message.
        validation.ErrorMessage = "Please select a color from the list";

        // Specify the validation area.
        CellArea area;
        area.StartRow = 0;
        area.EndRow = 4;
        area.StartColumn = 0;
        area.EndColumn = 0;

        // Add the validation area.
        validation.AreaList.Add(area);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "ListDataValidation.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "ListDataValidation.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();    
    }
}
