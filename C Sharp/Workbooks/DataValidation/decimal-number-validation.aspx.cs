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

public partial class DecimalNumberValidation : System.Web.UI.Page
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

        // Create a worksheet and get the first worksheet.
        Worksheet ExcelWorkSheet = workbook.Worksheets[0];

        // Obtain the existing Validations collection.
        ValidationCollection validations = ExcelWorkSheet.Validations;

        // Create a validation object adding to the collection list.
        Validation validation = validations[validations.Add()];

        // Set the validation type.
        validation.Type = ValidationType.Decimal;

        // Specify the operator.
        validation.Operator = OperatorType.Between;

        // Set the lower and upper limits.
        validation.Formula1 = Decimal.MinValue.ToString();

        validation.Formula2 = Decimal.MaxValue.ToString();

        // Set the error message.
        validation.ErrorMessage = "Please enter a valid integer or decimal number";

        // Specify the validation area of cells.
        CellArea area;
        area.StartRow = 0;
        area.EndRow = 9;
        area.StartColumn = 0;
        area.EndColumn = 0;

        // Add the area.
        validation.AreaList.Add(area);

        // Set the number formats to 2 decimal places for the validation area.        

        for (int i = 0; i < 10; i++)
        {
            Aspose.Cells.Style style = new Aspose.Cells.Style();
            style.Custom = "0.00";
           ExcelWorkSheet.Cells[i, 0].SetStyle(style);
        }

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "DecimalNumberValidation.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "DecimalNumberValidation.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();    
    }
}
