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

public partial class Workbooks_Data_UsingICustomFunction : System.Web.UI.Page
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
        //Open the workbook
        Workbook workbook = new Workbook();

        //Obtaining the reference of the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        //Adding a sample value to "A1" cell
        worksheet.Cells["B1"].PutValue(5);

        //Adding a sample value to "A2" cell
        worksheet.Cells["C1"].PutValue(100);

        //Adding a sample value to "A3" cell
        worksheet.Cells["C2"].PutValue(150);

        //Adding a sample value to "B1" cell
        worksheet.Cells["C3"].PutValue(60);

        //Adding a sample value to "B2" cell
        worksheet.Cells["C4"].PutValue(32);

        //Adding a sample value to "B2" cell
        worksheet.Cells["C5"].PutValue(62);

        //Adding custom formula to Cell A1
        workbook.Worksheets[0].Cells["A1"].Formula = "=MyFunc(B1,C1:C5)";

        //Calcualting Formulas
        workbook.CalculateFormula(false, new CustomFunction());

        //Assign resultant value to Cell A1
        workbook.Worksheets[0].Cells["A1"].PutValue(workbook.Worksheets[0].Cells["A1"].Value);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "UsingICustomFunction.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "UsingICustomFunction.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      

    }
}

public class CustomFunction : ICustomFunction
{

    public object CalculateCustomFunction(string functionName, System.Collections.ArrayList paramsList, System.Collections.ArrayList contextObjects)
    {
        //get value of first parameter
        decimal firstParamB1 = System.Convert.ToDecimal(paramsList[0]);

        //get value of second parameter
        Array secondParamC1C5 = (Array)(paramsList[1]);

        decimal total = 0M;

        // get every item value of second parameter
        foreach (object[] value in secondParamC1C5)
        {
            total += System.Convert.ToDecimal(value[0]);
        }

        total = total / firstParamB1;

        //return result of the function
        return total;
    }
}

