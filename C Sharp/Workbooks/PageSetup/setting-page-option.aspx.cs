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

public partial class Workbooks_PageSetup_SettingPageOption : System.Web.UI.Page
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

        Worksheet worksheet = workbook.Worksheets[0];
        //Set the orientation 
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        //You can either choose FitToPages or Zoom property but not both at the same time
        worksheet.PageSetup.Zoom = 10;
        //Set the paper size 
        worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
        //Set the print quality of the worksheet 
        worksheet.PageSetup.PrintQuality = 200;
        //Set the first page number of the worksheet pages
        worksheet.PageSetup.FirstPageNumber = 1;

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "SettingPageOption.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "SettingPageOption.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }
}
