using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using Aspose.Cells;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for Merge/UnMerge Cells
    /// </summary>
    public class MergeUnMergeCells : System.Web.UI.Page
    {
        protected System.Web.UI.WebControls.Button btnMerge;
        protected System.Web.UI.WebControls.Button btnUnMerge;

        private void Page_Load(object sender, System.EventArgs e)
        {
            // Put user code to initialize the page here
        }

        #region Web Form Designer generated code
        override protected void OnInit(EventArgs e)
        {
            //
            // CODEGEN: This call is required by the ASP.NET Web Form Designer.
            //
             if (Context != null && Context.Session != null)
            {
                InitializeComponent();
                base.OnInit(e);
            }
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnMerge.Click += new System.EventHandler(this.btnMerge_Click);
            this.btnUnMerge.Click += new System.EventHandler(this.btnUnMerge_Click);
            this.Load += new System.EventHandler(this.Page_Load);

        }
        #endregion

        private void btnMerge_Click(object sender, System.EventArgs e)
        {
            //Create a Workbook.
            Aspose.Cells.Workbook wbk = new Aspose.Cells.Workbook();

            //Create a Worksheet and get the first sheet.
            Aspose.Cells.Worksheet worksheet = wbk.Worksheets[0];

            //Create a Cells object ot fetch all the cells.
            Aspose.Cells.Cells cells = worksheet.Cells;

            //Merge some Cells (C6:E7) into a single C6 Cell.
            cells.Merge(5, 2, 2, 3);

            //Input data into C6 Cell.
            worksheet.Cells[5, 2].PutValue("This is my value");

            //Create a Style object to fetch the Style of C6 Cell.
            Aspose.Cells.Style style = worksheet.Cells[5, 2].GetStyle();

            //Create a Font object
            Aspose.Cells.Font font = style.Font;

            //Set the name.
            font.Name = "Times New Roman";

            //Set the font size.
            font.Size = 18;

            //Set the font color
            font.Color = Color.Blue;

            //Bold the text
            font.IsBold = true;

            //Make it italic
            font.IsItalic = true;

            //Set the backgrond color of C6 Cell to Red
            style.ForegroundColor = Color.Red;

            style.Pattern = BackgroundType.Solid;

            //Apply the Style to C6 Cell.
            cells[5, 2].SetStyle(style);

            //Save the excel file
            wbk.Save(HttpContext.Current.Response,"MergeCells.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));

			// End response to avoid unneeded html after xls
		    Response.End();
        }

        private void btnUnMerge_Click(object sender, System.EventArgs e)
        {
            //Create a Workbook.
            string path = System.Web.HttpContext.Current.Server.MapPath("~");
            path = path.Substring(0, path.LastIndexOf("\\"));
            path += @"\designer\Workbooks\MergeCells.xls";

            Aspose.Cells.Workbook wbk = new Aspose.Cells.Workbook(path);


            //Create a Worksheet and get the first sheet.
            Aspose.Cells.Worksheet worksheet = wbk.Worksheets[0];

            //Create a Cells object ot fetch all the cells.
            Aspose.Cells.Cells cells = worksheet.Cells;

            //Unmerge the cells.
            cells.UnMerge(5, 2, 2, 3);

            //Save the excel file
            wbk.Save(HttpContext.Current.Response, "UnMergeCells.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
	
			// End response to avoid unneeded html after xls
		    Response.End();
        }
    }
}


