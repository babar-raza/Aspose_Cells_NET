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
	/// Summary description for Column3D.
	/// </summary>
	public class Column3D : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.DropDownList ColumnType;
		protected System.Web.UI.WebControls.DropDownList WallsColor;
		protected System.Web.UI.WebControls.DropDownList FloorColor;
		protected System.Web.UI.WebControls.DropDownList Rotation;
		protected System.Web.UI.WebControls.DropDownList Elevation;
		protected System.Web.UI.WebControls.DropDownList DepthPercent;
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
            workbook.Save(HttpContext.Current.Response, "3DColumn." + ddlFileVersion.SelectedItem.Value.ToLower() , ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

            //Put string values in row cells of column 1
			cells["A1"].PutValue("Region");
			cells["A2"].PutValue("France");
			cells["A3"].PutValue("Germany");
			cells["A4"].PutValue("England");

            //Put values in row cells of column 3
			cells["B1"].PutValue("Marketing Costs");
			cells["B2"].PutValue(70000);
			cells["B3"].PutValue(55000);
			cells["B4"].PutValue(30000);
		}

		private void CreateCellsFormatting(Workbook workbook)
		{
            //Initialize Style1
			Style style1 = workbook.Styles[workbook.Styles.Add()];
			
            //Set border for Style1
			style1.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
			
            //set Font Property IsBold to true
            style1.Font.IsBold = true;

            //Set Alignment of Style
			style1.HorizontalAlignment = TextAlignmentType.Center;
			style1.VerticalAlignment = TextAlignmentType.Center;

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

			//Set the width of the specified column 
			cells.SetColumnWidth(1,15);

            //Apply Style to A1 and B1
			cells["A1"].SetStyle(style1);
			cells["B1"].SetStyle(style1);			
			
            //Initialize Style2
			Style style2 = workbook.Styles[workbook.Styles.Add()];

			//Copy data from another style object
			style2.Copy(style1);

            //Set Font IsBold Property to False
			style2.Font.IsBold = false;

            //Set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right;
			
            //Set foreground color
			style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
			
            //Set Style Pattern
            style2.Pattern = BackgroundType.Solid;
			
            //Apply Style to A2 and A4
            cells["A2"].SetStyle(style2);
			cells["A4"].SetStyle(style2);

            //Initialize Style3
			Style style3 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy the properties from Style2
            style3.Copy(style2);

			//Set cell format
			style3.Custom = "\"$\"#,##0";	
			
            //Apply Style to Cell B2 and B4
            cells["B2"].SetStyle(style3);
			cells["B4"].SetStyle(style3);			

            //initialize Style4
			Style style4 = workbook.Styles[workbook.Styles.Add()];

            //Copy the properties of Style2
			style4.Copy(style2);

			//Sets foreground color
			style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
			
            //Set Styte Pattern
            style4.Pattern = BackgroundType.Solid;	
			
            //Apply Style on cell A3
            cells["A3"].SetStyle(style4);


            //initialize Style5
			Style style5 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy the properties of Style4
            style5.Copy(style4);

			//Set cell format
			style5.Custom = "\"$\"#,##0";

            //Set Style on cell B3
			cells["B3"].SetStyle(style5);
		}	

		private void CreateStaticReport(Workbook workbook)
		{
            //Initialize Worksheet
			Worksheet sheet = workbook.Worksheets[0];
			
            //Set the name of the worksheet. 
			sheet.Name = "3D Column";
			
            //Set Gridlines invisible
            sheet.IsGridlinesVisible = false;		

			//Create chart
			int chartIndex = 0; 

            //Select Chart Type based on Values in Column Type drop down List
			switch ( ColumnType.SelectedItem.Text )
			{
				case "Column3D":						
					chartIndex = sheet.Charts.Add(ChartType.Column3D,5,1,29,10);
					break;					
				case "Column3DClustered":						
					chartIndex = sheet.Charts.Add(ChartType.Column3DClustered,5,1,29,10);
					break;
				case "Column3DStacked":
					chartIndex = sheet.Charts.Add(ChartType.Column3DStacked,5,1,29,10);
					break;
			}				

            //Initialize Chart
			Chart chart = sheet.Charts[chartIndex];	

			//Set properties of chart 
			chart.CategoryAxis.MajorGridLines.IsVisible = false;

			//Set properties of nseries
			chart.NSeries.Add("B2:B4", true);
			chart.NSeries.CategoryData = "A2:A4";
			chart.NSeries.IsColorVaried = true;
	
			//Set properties of chart title 
			chart.Title.Text = "Marketing Costs by Region";
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.Size = 12;

			//Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Region";
			chart.CategoryAxis.Title.TextFont.Color = Color.Black;
			chart.CategoryAxis.Title.TextFont.IsBold = true;
			chart.CategoryAxis.Title.TextHorizontalAlignment = TextAlignmentType.Center;
			chart.CategoryAxis.Title.TextFont.Size = 10;

			//Set properties of valueaxis title 
			chart.ValueAxis.Title.Text = "In Thousands";
			chart.ValueAxis.Title.TextFont.Color = Color.Black;
			chart.ValueAxis.Title.TextFont.IsBold = true;
			chart.ValueAxis.Title.TextFont.Size = 10;
			chart.ValueAxis.Title.RotationAngle = 90;				

			//Set the legend position  to Top
			chart.Legend.Position = LegendPositionType.Top;	

            //Set Borders of chart invisible
			chart.PlotArea.Border.IsVisible = false;

			//Set properties of chart based on values in controls from UI
			chart.Walls.ForegroundColor = Color.FromName(WallsColor.SelectedItem.Text);
			chart.Floor.ForegroundColor = Color.FromName(FloorColor.SelectedItem.Text);
			chart.RotationAngle = int.Parse(Rotation.SelectedItem.Text);
			chart.Elevation = int.Parse(Elevation.SelectedItem.Text);
			chart.DepthPercent = int.Parse(DepthPercent.SelectedItem.Text);	
		}        
	
	}
}
