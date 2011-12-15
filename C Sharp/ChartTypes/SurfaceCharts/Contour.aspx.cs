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
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;


namespace Aspose.Cells.Demos
{
	/// <summary>
	/// Summary description for Contour.
	/// </summary>
	public class Contour : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.DropDownList ChartTypeList;
		protected System.Web.UI.WebControls.Button btnProcess;
        protected System.Web.UI.WebControls.DropDownList ddlFileVersion;
	
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
			this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
			this.ID = "Contour";
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion

		private void btnProcess_Click(object sender, System.EventArgs e)
		{
            //Initialize Workbook
            Workbook workbook = new Workbook();

            //Set default font for workbook
            Style style = workbook.DefaultStyle;
            style.Font.Name = "Tahoma";
            workbook.DefaultStyle = style;

            //Insert Dummy Data
            CreateStaticData(workbook);

            //Apply Style on various cells
            CreateCellsFormatting(workbook);

            //Create Chart and Set Chart properties
            CreateStaticReport(workbook);

            //Create an object of SaveFormat
            SaveFormat saveFormat = new SaveFormat();

            //Check file format is xls
            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                //Set save format optoin to xls
                saveFormat = SaveFormat.Excel97To2003;
            }
            //Check file format is xlsx
            else if (ddlFileVersion.SelectedItem.Value == "XLSX")
            {
                //Set save format optoin to xlsx
                saveFormat = SaveFormat.Xlsx;
            }

            //Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "Contour." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
			Worksheet sheet = workbook.Worksheets[0];
			//Set the name of worksheet
			sheet.Name = "Data";
			sheet.IsGridlinesVisible = false;

			Cells cells = workbook.Worksheets[0].Cells;
			//Put a value into a cell
			cells["A1"].PutValue("Temperature");
			cells["A2"].PutValue("Seconds");
			cells["A3"].PutValue(0.2);
			cells["A4"].PutValue(0.3);
			cells["A5"].PutValue(0.4);
			cells["A6"].PutValue(0.5);
			cells["A7"].PutValue(0.6);
			cells["A8"].PutValue(0.7);
			cells["A9"].PutValue(0.8);
			cells["A10"].PutValue(0.9);
			cells["A11"].PutValue(1);			

			//Merge a specified range of cells into a single cell
			cells.Merge(0,1,2,1);
			cells["B1"].PutValue(10);
			cells["B3"].PutValue(99);
			cells["B4"].PutValue(107);
			cells["B5"].PutValue(119);
			cells["B6"].PutValue(135);
			cells["B7"].PutValue(155);
			cells["B8"].PutValue(184);
			cells["B9"].PutValue(193);
			cells["B10"].PutValue(295);
			cells["B11"].PutValue(384);

			//Merge a specified range of cells into a single cell
			cells.Merge(0,2,2,1);
			cells["C1"].PutValue(20);
			cells["C3"].PutValue(175);
			cells["C4"].PutValue(185);
			cells["C5"].PutValue(200);
			cells["C6"].PutValue(220);
			cells["C7"].PutValue(245);
			cells["C8"].PutValue(279);
			cells["C9"].PutValue(349);
			cells["C10"].PutValue(385);
			cells["C11"].PutValue(499);

			//Merge a specified range of cells into a single cell
			cells.Merge(0,3,2,1);
			cells["D1"].PutValue(30);
			cells["D3"].PutValue(250);
			cells["D4"].PutValue(260);
			cells["D5"].PutValue(275);
			cells["D6"].PutValue(275);
			cells["D7"].PutValue(320);
			cells["D8"].PutValue(356);
			cells["D9"].PutValue(392);
			cells["D10"].PutValue(405);
			cells["D11"].PutValue(459);

			//Merge a specified range of cells into a single cell
			cells.Merge(0,4,2,1);
			cells["E1"].PutValue(40);
			cells["E3"].PutValue(467);
			cells["E4"].PutValue(385);
			cells["E5"].PutValue(349);
			cells["E6"].PutValue(279);
			cells["E7"].PutValue(245);
			cells["E8"].PutValue(220);
			cells["E9"].PutValue(200);
			cells["E10"].PutValue(185);
			cells["E11"].PutValue(175);
			
			//Merge a specified range of cells into a single cell
			cells.Merge(0,5,2,1);
			cells["F1"].PutValue(50);
			cells["F3"].PutValue(400);
			cells["F4"].PutValue(305);
			cells["F5"].PutValue(209);
			cells["F6"].PutValue(192);
			cells["F7"].PutValue(163);
			cells["F8"].PutValue(144);
			cells["F9"].PutValue(118);
			cells["F10"].PutValue(59);
			cells["F11"].PutValue(25);
		}

		private void CreateCellsFormatting(Workbook workbook)
		{
            //Initialize Style
			Style style1 = workbook.Styles[workbook.Styles.Add()];
			//Set border setting of Style
			style1.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
			
            //Set Font to Bold
            style1.Font.IsBold = true;
			
            //Set alignmet of Style
            style1.HorizontalAlignment = TextAlignmentType.Center;

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

			//Set the width of the specified column 
			cells.SetColumnWidth(0,12);

            //Set Style for two rows
			cells["A1"].SetStyle(style1);
			cells["A2"].SetStyle(style1);
			cells["B1"].SetStyle(style1);
			cells["B2"].SetStyle(style1);
			cells["C1"].SetStyle(style1);
			cells["C2"].SetStyle(style1);
			cells["D1"].SetStyle(style1);
			cells["D2"].SetStyle(style1);
			cells["E1"].SetStyle(style1);
			cells["E2"].SetStyle(style1);
			cells["F1"].SetStyle(style1);
			cells["F2"].SetStyle(style1);

            //initialize Style2
			Style style2 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy data from another style object
			style2.Copy(style1);
			
            //Set Font to Normal
            style2.Font.IsBold = false;

            //Set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right;

			//Set foreground color
			style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
			style2.Pattern = BackgroundType.Solid;       
     
            //Loop over Cells and apply Style
			for ( int i = 2; i <= 10; i ++ )
			{
				if (i % 2 == 0)
				{
					cells[i,0].SetStyle(style2);
					cells[i,1].SetStyle(style2);
					cells[i,2].SetStyle(style2);
					cells[i,3].SetStyle(style2);
					cells[i,4].SetStyle(style2);
					cells[i,5].SetStyle(style2);
				}
			}
            //initialize Style3
			Style style3 = workbook.Styles[workbook.Styles.Add()];

            //Copy Style properties from another Object
			style3.Copy(style2);

			//Set foreground color
			style3.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
			
            //Set Pattern of Style
            style3.Pattern = BackgroundType.Solid;	

            //Loop over the cells and set style
			for ( int i = 2; i <= 10; i++)
			{
				if (i % 2 !=0)
				{
					cells[i, 0].SetStyle(style2);
					cells[i, 1].SetStyle(style3);
					cells[i,2].SetStyle(style3);
					cells[i,3].SetStyle(style3);
					cells[i,4].SetStyle(style3);
					cells[i,5].SetStyle(style3);
				}
			}
		}
		
		private void CreateStaticReport(Workbook workbook)
		{
			int sheetIndex = workbook.Worksheets.Add();
			Worksheet sheet = workbook.Worksheets[sheetIndex];
			//Set the name of worksheet
			sheet.Name = "Chart";

			//Create chart 
			int chartIndex = 0;
			switch ( ChartTypeList.SelectedItem.Text )
			{				
				case "SurfaceContour":
					chartIndex = sheet.Charts.Add(ChartType.SurfaceContour,1,1,25,10);
					break;
				case "SurfaceContourWireframe":
					chartIndex = sheet.Charts.Add(ChartType.SurfaceContourWireframe,1,1,25,10);
					break;
			}				
			Chart chart = sheet.Charts[chartIndex];
			
            //Set properties of chart
			chart.PlotArea.Border.IsVisible = false;
            		
			//Set properties of nseries
			chart.NSeries.Add("Data!B3:F11",true);
			chart.NSeries.CategoryData = "Data!A3:A11";
			chart.NSeries.IsColorVaried = true;
		
			Cells cells = workbook.Worksheets[0].Cells;

			for (int i = 0; i < chart.NSeries.Count; i ++)
			{				
				chart.NSeries[i].Name = cells[0,i + 1].StringValue;
			}		

			//Set properties of chart title
			chart.Title.Text = "Tensile strenth Measurements";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;
			chart.Title.TextHorizontalAlignment = TextAlignmentType.Center;
			chart.Title.TextVerticalAlignment = TextAlignmentType.Center;

			//Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Seconds";
			chart.CategoryAxis.Title.TextFont.Color = Color.Black;
			chart.CategoryAxis.Title.TextFont.IsBold = true;
			chart.CategoryAxis.Title.TextFont.Size = 10;
		}
		
	}
}
