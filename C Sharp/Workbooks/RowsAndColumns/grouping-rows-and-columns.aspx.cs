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
	/// Summary description for GroupingRowsAndColumns.
	/// </summary>
	public class GroupingRowsAndColumns : System.Web.UI.Page
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
            //Open template
            string path = MapPath("~");
            path = path.Substring(0, path.LastIndexOf("\\"));
            path += @"\designer\Workbooks\GroupingRowsAndColumns.xls";


            Workbook workbook = new Workbook(path);		

			GroupRowsAndColumns(workbook);

            workbook.Save(HttpContext.Current.Response, "GroupingRowsAndColumns.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
	
			// End response to avoid unneeded html after xls
		    Response.End();
		}

		private void Button2_Click(object sender, System.EventArgs e)
		{
            //Open template
            string path = MapPath("~");
            path = path.Substring(0, path.LastIndexOf("\\"));
            path += @"\designer\Workbooks\UnGroupingRowsAndColumns.xls";

			Workbook workbook = new Workbook(path);	
	
			UnGroupRowsAndColumns(workbook);

            workbook.Save(HttpContext.Current.Response, "UnGroupingRowsAndColumns.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
		}		

		private void GroupRowsAndColumns(Workbook workbook)
		{
			Worksheet worksheet = workbook.Worksheets[0];
			worksheet.Cells.GroupRows(0, 9);
			worksheet.Cells.GroupColumns(0, 1);

			//Set SummaryRowBelow property
			worksheet.Outline.SummaryRowBelow = true;

            //Set SummaryColumnRight property
			worksheet.Outline.SummaryColumnRight = false;
		}	
	
		private void UnGroupRowsAndColumns(Workbook workbook)
		{
			Worksheet worksheet = workbook.Worksheets[0];
				
			worksheet.Cells.UngroupRows(0, 9);
			worksheet.Cells.UngroupColumns(0, 1);	
		}
	}
}
