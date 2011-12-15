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
using Aspose.Cells.Pivot;

namespace Aspose.Cells.Demos
{
    public partial class Pivot_Table_MultiSource : System.Web.UI.Page
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

            path += @"\designer\Workbooks\PivotSource.xls";
            //Instantiating an Workbook object
            Workbook workbook = new Workbook(path);

            Worksheet sheet = workbook.Worksheets[0];

            PivotTableCollection pivotTables = sheet.PivotTables;
            
            String[] sourceData = new String[] { "=Sheet1!A1:C8", "=Sheet2!A1:C8" };
            PivotPageFields pageField = new PivotPageFields();
            String[] pageItems = new String[2];
            pageItems[0] = "Item1";
            pageItems[1] = "Item2";
            pageField.AddPageField(pageItems);
            pageItems = new String[2];
            pageItems[0] = "Item3";
            pageItems[1] = "Item4";
            pageField.AddPageField(pageItems);
            int[] TBPG = new int[2];

            TBPG[0] = 0;
            TBPG[1] = 1;
            
            //Sets which item label in each page field to use to identify the data range.
            pageField.AddIdentify(0, TBPG);
            TBPG = new int[2];
            TBPG[0] = 1;
            TBPG[1] = -1;
            pageField.AddIdentify(1, TBPG);
            int index = pivotTables.Add(sourceData, false, pageField, "E3", "PivotTable1");
            
            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "PivotTableMultipleSource.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "PivotTableMultipleSource.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End();   
        }
    }
}