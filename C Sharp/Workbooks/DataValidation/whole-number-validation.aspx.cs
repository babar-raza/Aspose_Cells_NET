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

public partial class WholeNumberValidation : System.Web.UI.Page
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
        //Initialize Workbook
        Workbook workbook = new Workbook();

        //Initialize WorkSheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Obtain the cells of the first worksheet.
        Cells cells = workbook.Worksheets[0].Cells;

        //Put a string value into A1 cell.
        cells["A1"].PutValue("Please enter whole number between 10 and 1000 only in this column.");

        //Accessing the Validations collection of the worksheet
        ValidationCollection validations = worksheet.Validations;

        //Creating a Validation object
        Validation validation = validations[validations.Add()];

        //Setting the validation type to whole number
        validation.Type = ValidationType.WholeNumber;

        //Setting the operator for validation to Between
        validation.Operator = OperatorType.Between;

        //Setting the minimum value for the validation
        validation.Formula1 = "10";

        //Setting the maximum value for the validation
        validation.Formula2 = "1000";

        validation.ErrorMessage = "Invalid Whole Number. Enter whole number between 10 and 1000 only.";

        //Applying the validation to a range of cells from A1 to B2 using the CellArea structure
        CellArea area;
        area.StartRow = 1;
        area.EndRow = 9;
        area.StartColumn = 0;
        area.EndColumn = 0;

        //Adding the cell area to Validation
        validation.AreaList.Add(area);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "WholeNumberValidation.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "WholeNumberValidation.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();    
    }
}
