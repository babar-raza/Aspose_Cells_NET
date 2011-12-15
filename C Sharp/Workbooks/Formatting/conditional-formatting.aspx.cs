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

public partial class Workbooks_Formatting_ConditionalFormatting : System.Web.UI.Page
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
        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        Worksheet sheet = workbook.Worksheets[0];

        //Adds an empty conditional formatting
        int index = sheet.ConditionalFormattings.Add();

        //Initialize FormatConditionCollection from newly inserted Index
        FormatConditionCollection fcs = sheet.ConditionalFormattings[index];

        //Sets the conditional format range.
        CellArea ca = new CellArea();
        ca.StartRow = 0;
        ca.EndRow = 0;
        ca.StartColumn = 0;
        ca.EndColumn = 0;

        //Assign FormatConditionCollection the Area
        fcs.AddArea(ca);


        //Adds condition.
        int conditionIndex = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "50", "100");

        //Sets the background color.
        FormatCondition fc = fcs[conditionIndex];

        //Set BackgroundColor
        fc.Style.BackgroundColor = Color.Red;



        //Adds an empty conditional formatting
        int index2 = sheet.ConditionalFormattings.Add();

        //Initialize FormatConditionCollection for newly added index
        FormatConditionCollection fcs2 = sheet.ConditionalFormattings[index2];

        //Sets the conditional format range.
        CellArea ca2 = new CellArea();
        ca2.StartRow = 2;
        ca2.EndRow = 2;
        ca2.StartColumn = 1;
        ca2.EndColumn = 1;

        //Assign FormatConditionCollection the Area
        fcs2.AddArea(ca2);

        //Adds condition.
        int conditionIndex2 = fcs2.AddCondition(FormatConditionType.Expression);
        
        //Sets the background color.
        FormatCondition fc2 = fcs2[conditionIndex2];

        //Set FormatCondition Object formula
        fc2.Formula1 = "=IF(SUM(B1:B2)>100,TRUE,FALSE)";

        //Set FormatCondition Object Background Color
        fc2.Style.BackgroundColor = Color.Red;

        sheet.Cells["B3"].Formula = "=SUM(B1:B2)";

        //Put Value in Cell C4 
        sheet.Cells["C4"].PutValue("If Sum of B1:B2 is greater than 100, B3 will have RED background");

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "ConditionalFormatting.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "ConditionalFormatting.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }

}



