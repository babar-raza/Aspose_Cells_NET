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
	/// Summary description for PercentStackedBar.
	/// </summary>
	public class PercentStackedBar : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.CheckBox CheckBoxShow3D;
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
            workbook.Save(HttpContext.Current.Response, "PercentStackedBar." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
            //Initialize worksheet
			Worksheet worksheet = workbook.Worksheets[0];
			
            //Set the name of worksheet
			worksheet.Name = "Data";
			
            //Set GridLines Invisible
            worksheet.IsGridlinesVisible = false;

            //initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;
			
            //put values for row 1
			cells["A1"].PutValue("Product Name");
			cells["B1"].PutValue("Quarter1");
			cells["C1"].PutValue("Quarter2");
			cells["D1"].PutValue("Quarter3");
			cells["E1"].PutValue("Quarter4");

            //put values for row 2
			cells["A2"].PutValue("Product1");
			cells["B2"].PutValue(0.33);
			cells["C2"].PutValue(0.21);
			cells["D2"].PutValue(0.35);
			cells["E2"].PutValue(0.22);

            //put values for row 3
			cells["A3"].PutValue("Product2");
			cells["B3"].PutValue(0.17);
			cells["C3"].PutValue(0.54);
			cells["D3"].PutValue(0.17);
			cells["E3"].PutValue(0.60);

            //put values for row 4
			cells["A4"].PutValue("Product3");
			cells["B4"].PutValue(0.50);
			cells["C4"].PutValue(0.25);
			cells["D4"].PutValue(0.48);
			cells["E4"].PutValue(0.18);
		}

		private void CreateCellsFormatting(Workbook workbook)
		{
            //Initialize Style1
			Style style1 = workbook.Styles[workbook.Styles.Add()];
			//Set borders
			style1.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
			style1.Font.IsBold = true;
			
            //Set Styles alignment
            style1.HorizontalAlignment = TextAlignmentType.Center;
			style1.VerticalAlignment = TextAlignmentType.Center;

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

			//Set the width of the specified column 
			cells.SetColumnWidth(0,12);

            //Set style of Row 1
			cells["A1"].SetStyle(style1);
			cells["B1"].SetStyle(style1);
			cells["C1"].SetStyle(style1);
			cells["D1"].SetStyle(style1);
			cells["E1"].SetStyle(style1);
			
            //initialize Style2
			Style style2 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy data from another style object
			style2.Copy(style1);
			
            //Set Font property IsBold to false
            style2.Font.IsBold = false;

            //set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right;
			
            //Set foreground color
			style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
			
            //Set Pattern of Style
            style2.Pattern = BackgroundType.Solid;
            
            //Set Style to cell A2 and A4
			cells["A2"].SetStyle(style2);
			cells["A4"].SetStyle(style2);

            //Initialize Style3
			Style style3 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy properties from Style2
            style3.Copy(style2);

			//Set cell format
			style3.Number = 9;
			
			//Loop over cells and Appky Style
            for( int i = 1 ; i <= 3; i ++ )
			{
				if ( i % 2 !=0)
				{
					cells[i, 1].SetStyle(style3);
					cells[i,2].SetStyle(style3);
					cells[i,3].SetStyle(style3);
					cells[i,4].SetStyle(style3);
				}
			}	

			//Initialize Style4
            Style style4 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy properties from Style2
            style4.Copy(style2);

			//Sets foreground color
			style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
			style4.Pattern = BackgroundType.Solid;

            //Set Style of cell A3
			cells["A3"].SetStyle(style4);

            //Initialize Style5
			Style style5 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy Style properties from Style4
            style5.Copy(style4);
			
            //Set cell format
			style5.Number = 9;

            //Loop over cells and Appky Style
			for ( int i = 1 ; i <= 3; i ++ )
			{
				if (i % 2 == 0)
				{
					cells[i,1].SetStyle(style5);
					cells[i,2].SetStyle(style5);
					cells[i,3].SetStyle(style5);
					cells[i,4].SetStyle(style5);
				}
			}	
		}		

		private void CreateStaticReport(Workbook workbook)
		{
            //Initialize worksheet of type Chart
			int sheetIndex = workbook.Worksheets.Add(SheetType.Chart);
			Worksheet sheet = workbook.Worksheets[sheetIndex];

			//Set the name of worksheet
			sheet.Name = "Chart";

            //Create chart, If Check box on Ui is Checked then create Bar3D100PercentStacked Chart else Bar100PercentStacked Chart
			int chartIndex = 0;	
			if (CheckBoxShow3D.Checked)
				chartIndex = sheet.Charts.Add(ChartType.Bar3D100PercentStacked,0,0,0,0);
			else
				chartIndex = sheet.Charts.Add(ChartType.Bar100PercentStacked,0,0,0,0);
			Chart chart = sheet.Charts[chartIndex];

			//Set properties of chart and hide gridLines based upon the state od check box on UI
			chart.CategoryAxis.MajorGridLines.IsVisible = false;			
			if (CheckBoxShow3D.Checked)
			  chart.PlotArea.Border.IsVisible = false;

			//Set properties of chart title
			chart.Title.Text = "Product contribution to total sales";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;

			//Set properties of nseries
			chart.NSeries.Add("Data!B2:E4", false);

            //Set Category Data from cells between B1 and E1
			chart.NSeries.CategoryData = "Data!B1:E1";

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

            //loop over the Chart's Nseries and Assign Name from Cell Values
			for ( int i = 0 ; i < chart.NSeries.Count ; i ++ )
			{				
				chart.NSeries[i].Name = cells[i+1,0].Value.ToString();
			}

			//Set properties of valueaxis title
			chart.ValueAxis.Title.Text = "% of total sales";
			chart.ValueAxis.Title.TextFont.Color = Color.Black;
			chart.ValueAxis.Title.TextFont.IsBold = true;
			chart.ValueAxis.Title.TextFont.Size = 10;
		}



	}
}
