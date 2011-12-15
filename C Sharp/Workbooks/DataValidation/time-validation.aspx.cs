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

public partial class TimeDataValidation : System.Web.UI.Page
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
        // Create a workbook.
        Workbook workbook = new Workbook();

        // Obtain the cells of the first worksheet.
        Cells cells = workbook.Worksheets[0].Cells;

        // Put a string value into A1 cell.
        cells["A1"].PutValue("Please enter Time b/w 09:00 and 11:30 'o Clock");

        // Wrap the text.
        cells["A1"].GetStyle().IsTextWrapped = true;
        //cells["A1"].Style.IsTextWrapped = true;

        // Set the row height and column width for the cells.
        cells.SetRowHeight(0, 31);

        cells.SetColumnWidth(0, 35);

        // Get the validations collection.
        ValidationCollection validations = workbook.Worksheets[0].Validations;

        // Add a new validation.
        Validation validation = validations[validations.Add()];

        // Set the data validation type.
        validation.Type = ValidationType.Time;

        // Set the operator for the data validation.
        validation.Operator = OperatorType.Between;

        // Set the value or expression associated with the data validation.
        validation.Formula1 = "09:00";

        // The value or expression associated with the second part of the data validation.
        validation.Formula2 = "11:30";

        // Enable the error.
        validation.ShowError = true;

        // Set the validation alert style.
        validation.AlertStyle = ValidationAlertType.Information;

        // Set the title of the data-validation error dialog box.
        validation.ErrorTitle = "Time Error";

        // Set the data validation error message.
        validation.ErrorMessage = "Enter a Valid Time";

        // Set and enable the data validation input message.
        validation.InputMessage = "Time Validation Type";

        validation.IgnoreBlank = true;

        validation.ShowInput = true;

        // Set a collection of CellArea which contains the data validation settings.
        CellArea cellArea;

        cellArea.StartRow = 0;

        cellArea.EndRow = 0;

        cellArea.StartColumn = 1;

        cellArea.EndColumn = 1;

        // Add the validation area.
        validation.AreaList.Add(cellArea);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "TimeValidation.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "TimeValidation.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();    
    }
}
