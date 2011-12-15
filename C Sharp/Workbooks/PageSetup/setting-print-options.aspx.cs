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

public partial class Workbooks_PageSetup_SettingPrintOptions : System.Web.UI.Page
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
        //Open template
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\book1.xls";


        Workbook workbook = new Workbook(path);


        PageSetup pageSetup = workbook.Worksheets[0].PageSetup;
        //Specify the cells range (from A1 cell to B2 cell) of the print area
        pageSetup.PrintArea = "A1:G5";

        //Define column numbers A & B as title columns
        pageSetup.PrintTitleColumns = "$A:$B";

        //Define row numbers 1 & 2 as title rows
        pageSetup.PrintTitleRows = "$1:$2";

        //Allow to print gridlines
        pageSetup.PrintGridlines = true;

        //Allow to print row/column headings
        pageSetup.PrintHeadings = true;

        //Allow to print worksheet in black & white mode
        pageSetup.BlackAndWhite = true;

        //Allow to print comments as displayed on worksheet
        pageSetup.PrintComments = PrintCommentsType.PrintInPlace;

        //Allow to print worksheet with draft quality
        pageSetup.PrintDraft = true;

        //Allow to print cell errors 
        pageSetup.PrintErrors = PrintErrorsType.PrintErrorsBlank;

        //Set the printing order of the pages to over then down
        pageSetup.Order = PrintOrderType.DownThenOver;

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "SettingPrintOptions.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "SettingPrintOptions.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }
}
