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

public partial class Workbooks_Data_FindOrSearchData : System.Web.UI.Page
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
        //Instantiate a new workbook
        Workbook workbook = new Workbook();

        //Set default font
        Aspose.Cells.Style style = workbook.DefaultStyle;
        style.Font.Name = "Tahoma";
        workbook.DefaultStyle = style;

        //Call Method to create data
        CreateSaticData(workbook);

        //Call Method to find data
        FindData(workbook);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "FindOrSearchData.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "FindOrSearchData.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      

    }

    private static void CreateSaticData(Workbook workbook)
    {
        //Get the cells collection in the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;
        
        //Put some values into cells
        cells["A1"].PutValue("Product ID");
        cells["A2"].PutValue(1);
        cells["A3"].PutValue(2);
        cells["A4"].PutValue(3);
        cells["A5"].PutValue(4);

        cells["A7"].PutValue(10);
        
        //Set a formula of the Cell. 
        cells["A7"].Formula = "=SUM(A2:A5)";

        cells["B1"].PutValue("Product Names");
        cells["B2"].PutValue("Apples");
        cells["B3"].PutValue("Bananas");
        cells["B4"].PutValue("Grapes");
        cells["B5"].PutValue("Oranges");
    }

    private static void FindData(Workbook workbook)
    {
        //Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        //Finds the cell with the input formula
        Aspose.Cells.Cell cell1 = sheet.Cells.FindFormula("=SUM(A2:A5)", null);

        //Find the cell with formla which contains the input string
        Aspose.Cells.Cell cell2 = sheet.Cells.FindFormulaContains("SUM", null);

        //Find the cell with the input integer or double
        Aspose.Cells.Cell cell3 = sheet.Cells.FindNumber(3, null);

        //Find the cell with the input string
        Aspose.Cells.Cell cell4 = sheet.Cells.FindString("Apples", null);

        //Find the cell containing with the input string
        Aspose.Cells.Cell cell5 = sheet.Cells.FindStringContains("anan", null);

        //Find the cell ending with the input string
        Aspose.Cells.Cell cell6 = sheet.Cells.FindStringEndsWith("as", null);

        //Find the cell starting with the input string
        Aspose.Cells.Cell cell7 = sheet.Cells.FindStringStartsWith("Gr", null);

        Cells cells = workbook.Worksheets[0].Cells;

        //Put some values into the cells
        cells["A9"].PutValue("Name of the cell with the input formula (=SUM(A2:A5)): " + cell1.Name);
        cells["A10"].PutValue("Name of the cell with formla which contains the input string (\"SUM\"): " + cell2.Name);
        cells["A11"].PutValue("Name of the cell with the input integer or double (3): " + cell3.Name);
        cells["A12"].PutValue("Name of the cell with the input string (\"Apples\"): " + cell4.Name);
        cells["A13"].PutValue("Name of the cell containing with the input string (\"anan\"): " + cell5.Name);
        cells["A14"].PutValue("Name of the cell ending with the input string (\"as\"): " + cell6.Name);
        cells["A15"].PutValue("Name of the cell starting with the input string (\"Gr\"): " + cell7.Name);
    }
}



