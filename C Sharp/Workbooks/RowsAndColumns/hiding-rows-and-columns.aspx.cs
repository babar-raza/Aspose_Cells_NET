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
	/// Summary description for HidingRowsAndColumns.
	/// </summary>
	public class HidingRowsAndColumns : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.Button Button1;
		protected System.Web.UI.WebControls.Button Button2;
	
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

			CreateStaticData(workbook);

			//Unhide the 3rd row and setting its height to 13.5
			sheet.Cells.UnhideRow(2, 13.5);
			//Unhide the 2nd column and setting its width to 15
			sheet.Cells.UnhideColumn(1, 15);

            workbook.Save(HttpContext.Current.Response, "DisplayRowsColumns.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
	
			// End response to avoid unneeded html after xls
		    Response.End();


		}

		private void Button2_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();

			Worksheet sheet = workbook.Worksheets[0];

			CreateStaticData(workbook);

			//Hide the 3rd row of the worksheet
			sheet.Cells.HideRow(2);
			//Hide the 2nd column of the worksheet
			sheet.Cells.HideColumn(1);

            workbook.Save(HttpContext.Current.Response, "HideRowsColumns.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
	
			// End response to avoid unneeded html after xls
		    Response.End();
		}	

		private void CreateStaticData(Workbook workbook)
		{
			//Set default font
			Style style = workbook.DefaultStyle;
			style.Font.Name = "Tahoma";
			workbook.DefaultStyle = style;

			Cells cells = workbook.Worksheets[0].Cells;			
			//Put a value into a cell
			cells["A1"].PutValue("Year");				
			cells["A2"].PutValue(2005);
			cells["A3"].PutValue(2006);

			cells["B1"].PutValue("No. of Employees");	
			cells["B2"].PutValue(98);
			cells["B3"].PutValue(113);	
		}
	

	}
}
