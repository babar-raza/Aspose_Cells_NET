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

public partial class Workbooks_Formatting_FontSetting : System.Web.UI.Page
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
        //Create a workbook
        Workbook workbook = new Workbook();
        
        //Get the cells collection in the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;
        
        //Put a value into the cell
        cells["A2"].PutValue("Aspose");
        
        //Get Style
        Aspose.Cells.Style style = cells["A2"].GetStyle();

        //Set the color of the font
        style.Font.Color = Color.Red;
        
        //Set Style
        cells["A2"].SetStyle(style);

        //Put a value into the cell
        cells["B2"].PutValue("Aspose");

        //Get Style
        style = cells["B2"].GetStyle();
        
        //Set a value indicating whether the font is bold
        style.Font.IsBold = true;

        //Set Style
        cells["B2"].SetStyle(style);
        
        //Put a value into the cell
        cells["C2"].PutValue("Aspose");

        //Get Style
        style = cells["C2"].GetStyle();
        
        //Set a value indicating whether the font is italic
        style.Font.IsItalic = true;

        //Set Style
        cells["C2"].SetStyle(style);
        
        //Put a value into the cell
        cells["A4"].PutValue("Aspose");

        //Get Style
        style = cells["A4"].GetStyle();
        
        //Set a value indicating whether the font is strikeout
        style.Font.IsStrikeout = true;

        //Set Style
        cells["A4"].SetStyle(style);
                
        //Put a value into the cell
        cells["B4"].PutValue("Aspose");

        //Get Style
        style = cells["B4"].GetStyle();
        
        //Set a value indicating whether the font is subscript
        style.Font.IsSubscript = true;

        //Set Style
        cells["B4"].SetStyle(style);
                
        //Put a value into the cell
        cells["C4"].PutValue("Aspose");

        //Get Style
        style = cells["C4"].GetStyle();
        
        //Set a value indicating whether the font is super script.
        style.Font.IsSuperscript = true;

        //Set Style
        cells["C4"].SetStyle(style);
      
        //Put a value into the cell
        cells["A6"].PutValue("Aspose");

        //Get Style
        style = cells["A6"].GetStyle();
        
        //Set the name of the font
        style.Font.Name = "Verdana";

        //Set Style
        cells["A6"].SetStyle(style);
        
        //Put a value into the cell
        cells["B6"].PutValue("Aspose");
       
        //Get Style
        style = cells["B6"].GetStyle();
      
        //Set the size of the font
        style.Font.Size = 15;

        //Set Style
        cells["B6"].SetStyle(style);

        //Put a value into the cell
        cells["C6"].PutValue("Aspose");

        //Get Style
        style = cells["C6"].GetStyle();
        
        //Set the font underline type
        style.Font.Underline = FontUnderlineType.Accounting;

        //Set Style
        cells["C6"].SetStyle(style);

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "FontSetting.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "FontSetting.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }
}




