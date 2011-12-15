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

public partial class Workbooks_Security_AdvancedProtection : System.Web.UI.Page
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


        //Instantiate a workbook
        Workbook workbook = new Workbook(path);



        //Get the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];
        //Get the protection in the sheet
        Protection protection = worksheet.Protection;

        //Restricting users to delete columns of the worksheet
        protection.AllowDeletingColumn = false;

        //Restricting users to delete row of the worksheet
        protection.AllowDeletingRow = false;

        //Restricting users to edit contents of the worksheet
        protection.AllowEditingContent = false;

        //Allowing users to edit objects of the worksheet
        protection.AllowEditingObject = true;

        //Allowing users to edit scenarios of the worksheet
        protection.AllowEditingScenario = true;

        //Restricting users to filter
        protection.AllowFiltering = false;

        //Allowing users to format cells of the worksheet
        protection.AllowFormattingCell = true;

        //Allowing users to format rows of the worksheet
        protection.AllowFormattingRow = true;

        //Allowing users to insert columns in the worksheet
        protection.AllowInsertingColumn = true;

        //Allowing users to insert hyperlinks in the worksheet
        protection.AllowInsertingHyperlink = true;

        //Allowing users to insert rows in the worksheet
        protection.AllowInsertingRow = true;

        //Allowing users to select locked cells of the worksheet
        protection.AllowSelectingLockedCell = true;

        //Allowing users to select unlocked cells of the worksheet
        protection.AllowSelectingUnlockedCell = true;

        //Allowing users to sort
        protection.AllowSorting = true;

        //Allowing users to use pivot tables in the worksheet
        protection.AllowUsingPivotTable = true;

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "AdvancedProtection.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "AdvancedProtection.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();
    }

}
