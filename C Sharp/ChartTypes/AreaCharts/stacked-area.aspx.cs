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
	/// Summary description for StackedArea.
	/// </summary>
	public class StackedArea : System.Web.UI.Page
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

        /// <summary>
        /// Create WorkBook, insert dummy data in worksheet
        /// Create chart based on the dummy data
        /// Save n xls or xlsx file format
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
            workbook.Save(HttpContext.Current.Response, "StackedArea." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
            // Initialize worksheet
			Worksheet sheet = workbook.Worksheets[0];
			
            //Set the name of worksheet
			sheet.Name = "Data";
			
            //Set Gridlines invisible
            sheet.IsGridlinesVisible = false;

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;
			
            //Put values for rows in column A
			cells["A1"].PutValue("Region");
			cells["A2"].PutValue("France");
			cells["A3"].PutValue("Germany");
			cells["A4"].PutValue("England");
				
			//Put values in cells for row 1
			cells["B1"].PutValue(2002);
			cells["C1"].PutValue(2003);
			cells["D1"].PutValue(2004);
			cells["E1"].PutValue(2005);
			cells["F1"].PutValue(2006);							

            //put values in cells for row 2
			cells["B2"].PutValue(5000);
			cells["C2"].PutValue(15000);
			cells["D2"].PutValue(35000);
			cells["E2"].PutValue(30000);
			cells["F2"].PutValue(20000);

            //put values in cells for row3
			cells["B3"].PutValue(10000);
			cells["C3"].PutValue(25000);
			cells["D3"].PutValue(40000);
			cells["E3"].PutValue(52000);
			cells["F3"].PutValue(60000);

            //put values in cells for row 4
			cells["B4"].PutValue(40000);
			cells["C4"].PutValue(45000);
			cells["D4"].PutValue(50000);
			cells["E4"].PutValue(55000);
			cells["F4"].PutValue(70000);
		}

		private void CreateCellsFormatting(Workbook workbook)
		{
            //initialize Style 1
			Style style1 = workbook.Styles[workbook.Styles.Add()];
			
            //Set border settings for style1
			style1.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
			
            //set Font property
            style1.Font.IsBold = true;
			
            //Set aligment settings for Style1
            style1.HorizontalAlignment = TextAlignmentType.Center;
			style1.VerticalAlignment = TextAlignmentType.Center;

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;
			
            //Apply Style1 on first row
            cells["A1"].SetStyle(style1);
			cells["B1"].SetStyle(style1);
			cells["C1"].SetStyle(style1);
			cells["D1"].SetStyle(style1);
			cells["E1"].SetStyle(style1);
			cells["F1"].SetStyle(style1);
			
            //Intialize Style2
			Style style2 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy data from another style object
			style2.Copy(style1);
			
            //Set font property isBold to false
			style2.Font.IsBold = false;
			
            //set Aligment properties of Style
            style2.HorizontalAlignment = TextAlignmentType.Right;
			style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
			
            //Set Pattern for Style2
            style2.Pattern = BackgroundType.Solid;	

            //Set Style to A2 and A4
			cells["A2"].SetStyle(style2);
			cells["A4"].SetStyle(style2);

            //intialize Style3
			Style style3 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy properties from Style2
            style3.Copy(style2);
			
            //Set cell format
			style3.Custom = "\"$\"#,##0";	
		
			
            //Loop over the cells and set Style
            for( int i = 1; i <= 3; i ++ )
			{
				if ( i % 2 !=0)
				{
					cells[i, 1].SetStyle(style3);
					cells[i,2].SetStyle(style3);
					cells[i,3].SetStyle(style3);
					cells[i,4].SetStyle(style3);
					cells[i,5].SetStyle(style3);
				}
			}	


            //initialize Style4
			Style style4 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy properties from style2
            style4.Copy(style2);

			//Sets foreground color
			style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
			
            //Set pattern for Style4
            style4.Pattern = BackgroundType.Solid;	

            //Apply Style4 to A3
			cells["A3"].SetStyle(style4);

            //Initialize Style5
			Style style5 = workbook.Styles[workbook.Styles.Add()];

            //Copy properties from Style4
			style5.Copy(style4);

			//Set cell format
			style5.Custom = "\"$\"#,##0";	
		
            //loop over the cells and Apply Style
			for ( int i = 1; i <= 3; i ++ )
			{
				if (i % 2 == 0)
				{	
					cells[i,1].SetStyle(style5);
					cells[i,2].SetStyle(style5);
					cells[i,3].SetStyle(style5);
					cells[i,4].SetStyle(style5);
					cells[i,5].SetStyle(style5);
				}
			}
		}		

		private void CreateStaticReport(Workbook workbook)
		{
            //get next index for Worksheet
			int sheetIndex = workbook.Worksheets.Add();

            //initialize worksheet on given index
			Worksheet sheet = workbook.Worksheets[sheetIndex];
			
            //Set the name of worksheet
			sheet.Name = "Chart";

            //Create chart, If Check box on Ui is Checked then create Area3DStacked Chart else AreaStacked
			int chartIndex = 0;				
			if ( CheckBoxShow3D.Checked )
				chartIndex = sheet.Charts.Add(ChartType.Area3DStacked,1,1,25,10);
			else
				chartIndex = sheet.Charts.Add(ChartType.AreaStacked,1,1,25,10);				  
			Chart chart = sheet.Charts[chartIndex];		
	
			//Set legend position to top
			chart.Legend.Position = LegendPositionType.Top;

            //if Check box on UI is checked then hide Gridlines else show them
			chart.CategoryAxis.MajorGridLines.IsVisible = false;
			if ( CheckBoxShow3D.Checked )
			 chart.PlotArea.Border.IsVisible = false;

			//Set properties of title
			chart.Title.Text = "Total Sales ";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;

			//Set properties of nseries
			chart.NSeries.Add("Data!B2:F4", false);

            //set category data from B1 to F1 Area
			chart.NSeries.CategoryData = "Data!B1:F1";

            //set visual properties for NSeries
			chart.NSeries.IsColorVaried = true;

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

            //loop over the Chart's Nseries and Assign Name from Cell Values
			for ( int i = 0; i < chart.NSeries.Count; i ++)
			{
				chart.NSeries[i].Name = cells[i+1,0].Value.ToString();
				chart.NSeries[i].Points[i].Area.ForegroundColor = Color.Red;
			}

			//Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Year(2002-2006)";
			chart.CategoryAxis.Title.TextFont.Color = Color.Black;
			chart.CategoryAxis.Title.TextFont.IsBold = true;
			chart.CategoryAxis.Title.TextFont.Size = 10;
			chart.CategoryAxis.AxisBetweenCategories = false;						
		}

      
		
	}
}
