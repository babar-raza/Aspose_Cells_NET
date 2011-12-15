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

public partial class TextLengthValidation : System.Web.UI.Page
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
        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Obtain the cells of the first worksheet.
        Cells cells = workbook.Worksheets[0].Cells;

        //Put a string value into A1 cell.
        cells["A1"].PutValue("Please enter a string not more than 5 chars");

        // Wrap the text.
        cells["A1"].GetStyle().IsTextWrapped = true;

        // Set row height and column width for the cell.
        cells.SetRowHeight(0, 31);

        cells.SetColumnWidth(0, 35);

        // Get the validations collection.
        ValidationCollection validations = workbook.Worksheets[0].Validations;

        // Add a new validation.
        Validation validation = validations[validations.Add()];

        // Set the data validation type.
        validation.Type = ValidationType.TextLength;

        // Set the operator for the data validation.
        validation.Operator = OperatorType.LessOrEqual;

        // Set the value or expression associated with the data validation.
        validation.Formula1 = "5";

        // Enable the error.
        validation.ShowError = true;

        // Set the validation alert style.
        validation.AlertStyle = ValidationAlertType.Warning;

        // Set the title of the data-validation error dialog box.
        validation.ErrorTitle = "Text Length Error";

        // Set the data validation error message.
        validation.ErrorMessage = "Your string is invalid because it has more than 5 characters. Enter valid string.";

        // Set and enable the data validation input message.
        validation.InputMessage = "TextLength Validation Type";

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
            workbook.Save(HttpContext.Current.Response, "TextLengthValidation.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "TextLengthValidation.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();    
    }
}
