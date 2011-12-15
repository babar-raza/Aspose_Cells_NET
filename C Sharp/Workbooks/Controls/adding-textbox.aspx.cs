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
using Aspose.Cells.Drawing;

public partial class Workbooks_Controls_AddTextbox : System.Web.UI.Page
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
        Workbook workbook = new Workbook();

        //Get the first worksheet in the book.
        Worksheet worksheet = workbook.Worksheets[0];

        //Add a new textbox to the collection.
        int textboxIndex = worksheet.TextBoxes.Add(2, 1, 160, 200);

        //Get the textbox object.
        Aspose.Cells.Drawing.TextBox textbox0 = worksheet.TextBoxes[textboxIndex];

        //Fill the text.
        textbox0.Text = "ASPOSE______The .NET & JAVA Component Publisher!";

        //Get the textbox text frame.
        Aspose.Cells.Drawing.MsoTextFrame textframe0 = textbox0.TextFrame;

        //Set the textbox to adjust it according to its contents.
        textframe0.AutoSize = true;

        //Set the placement.
        textbox0.Placement = Aspose.Cells.Drawing.PlacementType.FreeFloating;

        //Set the font color.
        textbox0.Font.Color = System.Drawing.Color.Blue;

        //Set the font to bold.
        textbox0.Font.IsBold = true;

        //Set the font size.
        textbox0.Font.Size = 14;

        //Set font attribute to italic.
        textbox0.Font.IsItalic = true;

        //Add a hyperlink to the textbox.
        textbox0.AddHyperlink("http://www.aspose.com/");

        //Get the filformat of the textbox.
        MsoFillFormat fillformat = textbox0.FillFormat;

        //Set the fillcolor.
        fillformat.ForeColor = System.Drawing.Color.Silver;

        //Get the lineformat type of the textbox.
        MsoLineFormat lineformat = textbox0.LineFormat;

        //Set the line style.
        lineformat.Style = MsoLineStyle.ThinThick;

        //Set the line weight.
        lineformat.Weight = 6;

        //Set the dash style to squaredot.
        lineformat.DashStyle = MsoLineDashStyle.SquareDot;

        //Add another textbox.
        textboxIndex = worksheet.TextBoxes.Add(15, 4, 85, 120);

        //Get the second textbox.
        Aspose.Cells.Drawing.TextBox textbox1 = worksheet.TextBoxes[textboxIndex];

        //Input some text to it.
        textbox1.Text = "This is another simple text box";

        //Set the placement type as the textbox will move and resize with cells.
        textbox1.Placement = Aspose.Cells.Drawing.PlacementType.MoveAndSize; 

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "TextBox.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "TextBox.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();
    }

}



