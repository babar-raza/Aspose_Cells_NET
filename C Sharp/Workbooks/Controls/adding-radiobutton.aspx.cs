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
using Aspose.Cells.Drawing;

public partial class Workbooks_Controls_AddRadioButton : System.Web.UI.Page
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
        //Instantiate a new Workbook.
        Workbook excelbook = new Workbook();

        //Insert a value in C2 cell
        excelbook.Worksheets[0].Cells["C2"].PutValue("Age Groups");

        //Get style from C2 cell
        Aspose.Cells.Style style = excelbook.Worksheets[0].Cells["C2"].GetStyle();

        //Set the font text bold.
        style.Font.IsBold = true;

        //Set style to C2 Cell
        excelbook.Worksheets[0].Cells["C2"].SetStyle(style);

        //Add a radio button to the first sheet.
        Aspose.Cells.Drawing.RadioButton radio1 = excelbook.Worksheets[0].Shapes.AddRadioButton(3, 0, 2, 0, 30, 110);

        //Set its text string.
        radio1.Text = "20-29";

        //Set A1 cell as a linked cell for the radio button.
        radio1.LinkedCell = "A1";

        //Make the radio button 3-D.
        radio1.Shadow = true;

        //Set the foreground color of the radio button.
        radio1.FillFormat.ForeColor = Color.LightGreen;

        // set the line style of the radio button.
        radio1.LineFormat.Style = MsoLineStyle.ThickThin;

        //Set the weight of the radio button.
        radio1.LineFormat.Weight = 4;

        //Set the line color of the radio button.
        radio1.LineFormat.ForeColor = Color.Blue;

        //Set the dash style of the radio button.
        radio1.LineFormat.DashStyle = MsoLineDashStyle.Solid;

        //Make the line format visible.
        radio1.LineFormat.IsVisible = true;

        //Make the fill format visible.
        radio1.FillFormat.IsVisible = true;

        //Add another radio button to the first sheet.
        Aspose.Cells.Drawing.RadioButton radio2 = excelbook.Worksheets[0].Shapes.AddRadioButton(6, 0, 2, 0, 30, 110);

        //Set its text string.
        radio2.Text = "30-39";

        //Set A1 cell as a linked cell for the radio button.
        radio2.LinkedCell = "A1";

        //Make the radio button 3-D.
        radio2.Shadow = true;

        //Set the foreground color of the radio button.
        radio2.FillFormat.ForeColor = Color.LightGreen;

        // set the line style of the radio button.
        radio2.LineFormat.Style = MsoLineStyle.ThickThin;

        //Set the weight of the radio button.
        radio2.LineFormat.Weight = 4;

        //Set the line color of the radio button.
        radio2.LineFormat.ForeColor = Color.Blue;

        //Set the dash style of the radio button.
        radio2.LineFormat.DashStyle = MsoLineDashStyle.Solid;

        //Make the line format visible.
        radio2.LineFormat.IsVisible = true;

        //Make the fill format visible.
        radio2.FillFormat.IsVisible = true;

        //Add another radio button to the first sheet.
        Aspose.Cells.Drawing.RadioButton radio3 = excelbook.Worksheets[0].Shapes.AddRadioButton(9, 0, 2, 0, 30, 110);

        //Set its text string.
        radio3.Text = "40-49";

        //Set A1 cell as a linked cell for the radio button.
        radio3.LinkedCell = "A1";

        //Make the radio button 3-D.
        radio3.Shadow = true;

        //Set the foreground color of the radio button.
        radio3.FillFormat.ForeColor = Color.LightGreen;

        // set the line style of the radio button.
        radio3.LineFormat.Style = MsoLineStyle.ThickThin;

        //Set the weight of the radio button.
        radio3.LineFormat.Weight = 4;

        //Set the line color of the radio button.
        radio3.LineFormat.ForeColor = Color.Blue;

        //Set the dash style of the radio button.
        radio3.LineFormat.DashStyle = MsoLineDashStyle.Solid;

        //Make the line format visible.
        radio3.LineFormat.IsVisible = true;

        //Make the fill format visible.
        radio3.FillFormat.IsVisible = true;

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            excelbook.Save(HttpContext.Current.Response, "ComboBox.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            excelbook.Save(HttpContext.Current.Response, "ComboBox.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        // End response to avoid unneeded html after xls
        System.Web.HttpContext.Current.Response.End();
    }

}



