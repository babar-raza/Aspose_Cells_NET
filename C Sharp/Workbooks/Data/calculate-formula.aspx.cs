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

public partial class Workbooks_Data_CalculateFormula : System.Web.UI.Page
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
        //Open template
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\Workbooks\CalculateFormula.xls";

        //Instantiate a workbook
        Workbook workbook = new Workbook(path);

        //Get the cells collection in the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;
        for (int i = 11; i < 86; i++)
        {
            //Get a string value from a cell
            string strFormula = cells[i, 2].StringValue;
            //Set a formula of the Cell
            cells[i, 3].Formula = strFormula;
        }

        //Calculates the result of formulas
        workbook.CalculateFormula();
        for (int i = 11; i < 86; i++)
        {
            //Put values obtaining the calculated values
            cells[i, 4].PutValue(cells[i, 3].Value);
        }

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "CalculateFormula.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "CalculateFormula.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      

    }
}

