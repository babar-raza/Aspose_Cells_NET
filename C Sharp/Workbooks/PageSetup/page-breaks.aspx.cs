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
	/// Summary description for PageBreaks.
	/// </summary>
	public class PageBreaks : System.Web.UI.Page
	{
        protected System.Web.UI.WebControls.DropDownList ddlFileVersion;
		protected System.Web.UI.WebControls.Button Button2;
		protected System.Web.UI.WebControls.Button Button3;
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
			this.Button3.Click += new System.EventHandler(this.Button3_Click);
			this.Button2.Click += new System.EventHandler(this.Button2_Click);
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion

		private void Button1_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();
			Worksheet sheet = workbook.Worksheets[0];

			CreateStaticData(workbook);	
			
			//Add a page break at cell B2
			sheet.HorizontalPageBreaks.Add("B2");			
			sheet.VerticalPageBreaks.Add("B2");

            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "AddPageBreaks.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "AddPageBreaks.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End();  
		}

		private void CreateStaticData(Workbook workbook)
		{
			Cells cells = workbook.Worksheets[0].Cells;
			//Put a value into a cell
			cells["A1"].PutValue("World");
			cells["A2"].PutValue("Aspose");
			cells["A3"].PutValue(100);
			cells["B1"].PutValue(200);
			cells["B2"].PutValue(300);
			cells["B3"].PutValue(500);
		}

		private void Button2_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();
			Worksheet sheet = workbook.Worksheets[0];

			CreateStaticData(workbook);
			
			//Add a page break at cell B2
			sheet.HorizontalPageBreaks.Add("B2");			
			sheet.VerticalPageBreaks.Add("B2");
			sheet.HorizontalPageBreaks.Add(5, 1);
			sheet.HorizontalPageBreaks.Add(6, 1, 10);
			//Remove a page break at cell 
			sheet.HorizontalPageBreaks.RemoveAt(0);
			sheet.VerticalPageBreaks.RemoveAt(0);

            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "RemovetPageBreaks.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "RemovetPageBreaks.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End();  

		}

		private void Button3_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();			

			CreateStaticData(workbook);

			//Clear all page breaks
			workbook.Worksheets[0].HorizontalPageBreaks.Clear();
			workbook.Worksheets[0].VerticalPageBreaks.Clear();

            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "ClearPageBreaks.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "ClearPageBreaks.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End();  
		}
	}
}
