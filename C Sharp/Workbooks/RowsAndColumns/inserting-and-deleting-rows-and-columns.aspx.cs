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
	/// Summary description for InsertingAndDeletingRowsAndColumns.
	/// </summary>
	public class InsertingAndDeletingRowsAndColumns : System.Web.UI.Page
	{
        protected System.Web.UI.WebControls.DropDownList ddlFileVersion;
		protected System.Web.UI.WebControls.Button Button2;
		protected System.Web.UI.WebControls.Button Button1;
	
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
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{    
			this.Button1.Click += new System.EventHandler(this.Button1_Click);
			this.Button2.Click += new System.EventHandler(this.Button2_Click);
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion

		private void Button1_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();
			Worksheet sheet = workbook.Worksheets[0];

            Cells cells = workbook.Worksheets[0].Cells;

            //Put values into a cell
            cells["A1"].PutValue("1st Row & Column");
            cells["A2"].PutValue("2nd Row");
            cells["A3"].PutValue("3rd Row");
            cells["A4"].PutValue("4th Row");
            cells["A5"].PutValue("5th Row");
            cells["A6"].PutValue("6th Row");
            cells["A7"].PutValue("7th Row");
            cells["A8"].PutValue("8th Row");
            cells["A9"].PutValue("9th Row");
            cells["A10"].PutValue("10th Row");
            cells["A11"].PutValue("11th Row");
            cells["A12"].PutValue("12th Row");
            cells["A13"].PutValue("13th Row");
            cells["A14"].PutValue("14th Row");

            cells["B1"].PutValue("2nd Column");
            cells["C1"].PutValue("3rd Column");
            cells["D1"].PutValue("4th Column");
            cells["E1"].PutValue("5th Column");
            
            sheet.AutoFitColumns();

            //Insert 10 rows from the 3rd row
			sheet.Cells.InsertRows(2, 10);

			//Insert 3rd column 
			sheet.Cells.InsertColumn(2);

            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "InsertRowsAndColumns.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "InsertRowsAndColumns.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End();      
		}

		private void Button2_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();
			Worksheet sheet = workbook.Worksheets[0];

			Cells cells = workbook.Worksheets[0].Cells;
			//Put a value into a cell
			cells["A1"].PutValue("1st Row & Column");
			cells["A2"].PutValue("2nd Row");
            cells["A3"].PutValue("3rd Row");
            cells["A4"].PutValue("4th Row");
            cells["A5"].PutValue("5th Row");
            cells["A6"].PutValue("6th Row");
            cells["A7"].PutValue("7th Row");
            cells["A8"].PutValue("8th Row");
            cells["A9"].PutValue("9th Row");
            cells["A10"].PutValue("10th Row");
            cells["A11"].PutValue("11th Row");
            cells["A12"].PutValue("12th Row");
            cells["A13"].PutValue("13th Row");
            cells["A14"].PutValue("14th Row");

            cells["B1"].PutValue("2nd Column");
            cells["C1"].PutValue("3rd Column");
            cells["D1"].PutValue("4th Column");
            cells["E1"].PutValue("5th Column");

            sheet.AutoFitColumns();

			//Delete 10 rows from the 3rd row
			sheet.Cells.DeleteRows(2,10);
			
            //Delete 3rd column
			sheet.Cells.DeleteColumn(2);

            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "DeleteRowsAndColumns.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "DeleteRowsAndColumns.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End();      
		}
	}
}
