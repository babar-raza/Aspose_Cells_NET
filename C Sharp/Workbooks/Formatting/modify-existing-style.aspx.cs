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

public partial class Workbooks_Formatting_ModifyExistingStyle : System.Web.UI.Page
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
        //Create a workbook.
        Workbook workbook = new Workbook();

        //Create a new style object.
        Aspose.Cells.Style style = workbook.Styles[workbook.Styles.Add()];

        //Set the number format.
        style.Number = 14;

        //Set the font color to red color.
        style.Font.Color = System.Drawing.Color.Red;

        //Name the style.
        style.Name = "Style1";

        //Get the first worksheet cells.
        Cells cells = workbook.Worksheets[0].Cells;
        
        //Put value in cell
        cells["A1"].PutValue("Original Color Red & Modified Color Blue");

        //Specify the style (described above) to A1 cell.
        cells["A1"].SetStyle(style);

        //Create a range (B1:D1).
        Range range = cells.CreateRange("B1", "D1");

        //Initialize styleflag struct.
        StyleFlag flag = new StyleFlag();

        //Set all formatting attributes on.
        flag.All = true;

        //Apply the style (described above)to the range.
        range.ApplyStyle(style, flag);

        //Modify the style (described above) and change the font color from red to blue.
        style.Font.Color = System.Drawing.Color.Blue;

        //Done! Since the named style (described above) has been set to a cell and range, 
        //the change would be Reflected(new modification is implemented) to cell(A1) and //range (B1:D1).
        style.Update();

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "ModifyExistingStyle.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "ModifyExistingStyle.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      
    }

}



