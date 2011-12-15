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
using System.IO;

public partial class Workbooks_DrawingObjects_AddingWordArt : System.Web.UI.Page
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
        //initialize the workbook object
        Workbook workbook = new Workbook();
        
        //Get first worksheet in the workbook
        Worksheet sheet = workbook.Worksheets[0];

        //Apply WordArt Style with font settings
        sheet.Shapes.AddTextEffect(Aspose.Cells.Drawing.MsoPresetTextEffect.TextEffect1, "Aspose.Cells for .NET", "Arial", 15, true, true, 5, 5, 2, 2, 100, 175);

        //Apply WordArt Style with font settings
        sheet.Shapes.AddTextEffect(Aspose.Cells.Drawing.MsoPresetTextEffect.TextEffect2, "Aspose.Cells for Java", "Verdana", 30, true, false, 10, 5, 2, 2, 100, 100);

        //Apply WordArt Style with font settings
        sheet.Shapes.AddTextEffect(Aspose.Cells.Drawing.MsoPresetTextEffect.TextEffect3, "Aspose.Cells for Reporting Services", "Times New Roman", 25, false, true, 15, 5, 2, 2, 100, 150);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "AddingWordArt.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "AddingWordArt.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();  
    }
}
