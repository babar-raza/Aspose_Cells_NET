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

public partial class Workbooks_Data_HelloWorld : System.Web.UI.Page
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
        //Open template
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\Workbooks\HelloWorld.xls";

        //Create a workbook object
        Workbook workbook = new Workbook(path);
  

        //Get the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        //Get the cells collection in the sheet
        Cells cells = worksheet.Cells;

        //Put a string value into the cell using its name
        cells["A1"].PutValue("Cell Value");

        //put a string value into the cell using its name
        cells["A2"].PutValue("Hello World");

        //Put an boolean value into the cell using its name
        cells["A3"].PutValue(true);

        //Put an int value into the cell using its name
        cells["A4"].PutValue(100);

        //Put an double value into the cell using its name
        cells["A5"].PutValue(2856.5);

        //Put an string value that can be converted to other data type if appropriate
        cells["A6"].PutValue((123.6).ToString(), true);

        //Put an object value into the cell using its name
        object obj = "Aspose";
        cells["A7"].PutValue(obj);

        //Put an datetime value into the cell
        DateTime dt = DateTime.Now;
        cells["A8"].PutValue(dt);
        Aspose.Cells.Style style = cells["A8"].GetStyle();
        style.Font.IsBold = true;
        cells["A8"].SetStyle(style);
        
        //Put a string value into the cell using its row and column
        cells[0, 1].PutValue("Cell Value Type");

        for (int i = 1; i < 8; i++)
        {
            switch (cells[i, 0].Type)
            {
                //Cell value is boolean
                case CellValueType.IsBool:
                    cells[i, 1].PutValue("IsBool");
                    break;
                //Cell value is datetime
                case CellValueType.IsDateTime:
                    cells[i, 1].PutValue("IsDateTime");
                    break;
                //Blank cell
                case CellValueType.IsNull:
                    cells[i, 1].PutValue("IsNull");
                    break;
                //Cell value is numeric
                case CellValueType.IsNumeric:
                    cells[i, 1].PutValue("IsNumeric");
                    break;
                //Cell value is string
                case CellValueType.IsString:
                    cells[i, 1].PutValue("IsString");
                    break;
                //Cell value type is unknown
                case CellValueType.IsUnknown:
                    cells[i, 1].PutValue("IsUnknown");
                    break;
            }
        }

        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "HelloWorld.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "HelloWorld.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      

    }

}
