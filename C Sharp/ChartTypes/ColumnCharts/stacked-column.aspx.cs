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
	/// Summary description for StackedColumn.
	/// </summary>
	public class StackedColumn : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.CheckBox checkBoxShow3D;
		protected System.Web.UI.WebControls.Button btnProcess;
        protected System.Web.UI.WebControls.DropDownList ddlFileVersion;
	
		private void Page_Load(object sender, System.EventArgs e)
		{
			// Put user code to initialize the page here			
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
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
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion

        protected void btnProcess_Click(object sender, EventArgs e)
        {            
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
            CreateStaticReport(workbook); ;

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
            workbook.Save(HttpContext.Current.Response, "PercentStackedColumn." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
            // note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
            //Initialize Worksheet
			Worksheet sheet = workbook.Worksheets[0];

			//Set the name of worksheet
			sheet.Name = "Data";

            //Set worksheets Gridlines to invisible
			sheet.IsGridlinesVisible = false;			


            //initialize cells
			Cells cells = workbook.Worksheets[0].Cells;				

			//Put values in rows for Column1
			cells["A1"].PutValue("Year");
			cells["A2"].PutValue(2004);
			cells["A3"].PutValue(2005);
			cells["A4"].PutValue(2006);

            //Put values in rows for Column2
            cells["B2"].PutValue(20000);
			cells["B3"].PutValue(40000);
			cells["B4"].PutValue(40000);

            //Put values in rows for Column3
			cells["C2"].PutValue(30000);
			cells["C3"].PutValue(20000);
			cells["C4"].PutValue(50000);
			
            //Put value in CElls B1 and B2 for ROw Headers
			cells["B1"].PutValue("Product1"); 
			cells["C1"].PutValue("Product2");
		}

		private void CreateCellsFormatting(Workbook workbook)
		{
            //initialize Style1
			Style style1 = workbook.Styles[workbook.Styles.Add()];
			//Set Boder settings for style1
			style1.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
			
            //Set Font Style
            style1.Font.IsBold = true;

            //Set Alignments for Style2
			style1.HorizontalAlignment = TextAlignmentType.Center;
			style1.VerticalAlignment = TextAlignmentType.Center;

            ////initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

            //Apply Style to Row Headers
			cells["A1"].SetStyle(style1);
			cells["B1"].SetStyle(style1);
			cells["C1"].SetStyle(style1);

            //initialize Style2
			Style style2 = workbook.Styles[workbook.Styles.Add()];

			//Copy data from another style object
			style2.Copy(style1);	
			style2.Font.IsBold = false;
			style2.HorizontalAlignment = TextAlignmentType.Right;
			
            //Set foreground color
			style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
			
            //Set Style Pattern
            style2.Pattern = BackgroundType.Solid;	

			//Set Style to A2 and A4
            cells["A2"].SetStyle(style2);
			cells["A4"].SetStyle(style2);

            //initialize Style3
			Style style3 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style3.Copy(style2);
			
            //Set cell format
			style3.Custom = "\"$\"#,##0";

            //Set Style to Cells B, C2, B4 and C4
			cells["B2"].SetStyle(style3);
			cells["C2"].SetStyle(style3);
			cells["B4"].SetStyle(style3);
			cells["C4"].SetStyle(style3);

            //initialize Style4
			Style style4 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
			style4.Copy(style2);

			//Sets foreground color
			style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
			
            //Set Style Pattern
            style4.Pattern = BackgroundType.Solid;	

            //Set Style to cell A3
			cells["A3"].SetStyle(style4);

            //initialize Style5
			Style style5 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style5.Copy(style4);
			
            //Set cell format
			style5.Custom = "\"$\"#,##0";

            //Set Style to Cells B3 and C3
			cells["B3"].SetStyle(style5);
			cells["C3"].SetStyle(style5);
		}		

		private void CreateStaticReport(Workbook workbook)
		{
            //get next index for new worksheet
			int sheetIndex = workbook.Worksheets.Add();

            //Initalize worksheet for given index
			Worksheet sheet = workbook.Worksheets[sheetIndex];
			
            //Set the name of worksheet
			sheet.Name = "Chart";									

			//Create Chart
			int chartIndex = 0;
			// Show as 2d or 3d depending on the state of check Box on UI
			if(checkBoxShow3D.Checked)
				chartIndex = sheet.Charts.Add(ChartType.Column3DStacked, 1, 1, 25, 10);
			else
				chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 1, 1, 25, 10);
			Chart chart = sheet.Charts[chartIndex];

			//Set properies to chart
			chart.CategoryAxis.MajorGridLines.IsVisible =false;
			if (checkBoxShow3D.Checked)
				chart.PlotArea.Border.IsVisible = false;

			//Set properies to chart title
			chart.Title.Text = "Product  Sales";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;

			//Set properies to nseries
			chart.NSeries.CategoryData = "Data!A2:A4";

            //Set NSeries Data
			chart.NSeries.Add("Data!B2:C4", true);	
	
            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;
			

            //Iterate over the NSeries and assign it name from values cells
			for( int i = 0 ; i < chart.NSeries.Count; i ++ )
            {
				chart.NSeries[i].Name = cells[0,i + 1].Value.ToString();
			}
			  
			//Set properies to categoryaxis
			chart.CategoryAxis.Title.Text = "Year(2004-2006)";
			chart.CategoryAxis.Title.TextFont.Color = Color.Black;
			chart.CategoryAxis.Title.TextFont.Size = 10;
			chart.CategoryAxis.Title.TextFont.IsBold = true;			
		
			//Set the legend position To Tip
			chart.Legend.Position = LegendPositionType.Top;	
		}        		
	}
}
