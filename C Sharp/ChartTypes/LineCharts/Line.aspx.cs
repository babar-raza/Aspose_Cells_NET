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
	/// Summary description for Line.
	/// </summary>
	public class Line : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.DropDownList ChartTypeList;
		protected System.Web.UI.WebControls.DropDownList NSeriesMarkerStyle;
		protected System.Web.UI.WebControls.DropDownList NMarkBackColor;
		protected System.Web.UI.WebControls.DropDownList NMarkForeColor;
		protected System.Web.UI.WebControls.DropDownList NSeriesMarkSize;
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
            workbook.Save(HttpContext.Current.Response, "Line." + ddlFileVersion.SelectedItem.Value.ToLower() , ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

        private void CreateStaticData(Workbook workbook)
        {
            //Initialize Cells object
            Cells cells = workbook.Worksheets[0].Cells;

            //Put string into a cells of Column A
            cells["A1"].PutValue("Region");
            cells["A2"].PutValue("France");
            cells["A3"].PutValue("Germany");
            cells["A4"].PutValue("England");

            //Put a value into a Row 1
            cells["B1"].PutValue(2002);
            cells["C1"].PutValue(2003);
            cells["D1"].PutValue(2004);
            cells["E1"].PutValue(2005);
            cells["F1"].PutValue(2006);

            //Put a value into a Row 2
            cells["B2"].PutValue(40000);
            cells["C2"].PutValue(45000);
            cells["D2"].PutValue(50000);
            cells["E2"].PutValue(55000);
            cells["F2"].PutValue(70000);

            //Put a value into a Row 3
            cells["B3"].PutValue(10000);
            cells["C3"].PutValue(25000);
            cells["D3"].PutValue(40000);
            cells["E3"].PutValue(52000);
            cells["F3"].PutValue(60000);

            //Put a value into a Row 4
            cells["B4"].PutValue(5000);
            cells["C4"].PutValue(15000);
            cells["D4"].PutValue(35000);
            cells["E4"].PutValue(30000);
            cells["F4"].PutValue(20000);
        }

        private void CreateCellsFormatting(Workbook workbook)
        {
            //Initialize Style Object
            Style style1 = workbook.Styles[workbook.Styles.Add()];

            //Set borders setting for Style
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

            //Set Alignments for Style
            style1.HorizontalAlignment = TextAlignmentType.Center;
            style1.VerticalAlignment = TextAlignmentType.Center;

            //Initalize Cells Object
            Cells cells = workbook.Worksheets[0].Cells;

            //Set style for Row 1
            cells["A1"].SetStyle(style1);
            cells["B1"].SetStyle(style1);
            cells["C1"].SetStyle(style1);
            cells["D1"].SetStyle(style1);
            cells["E1"].SetStyle(style1);
            cells["F1"].SetStyle(style1);

            // Initialize Style 2
            Style style2 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style2.Copy(style1);

            //Set IsBold Off
            style2.Font.IsBold = false;

            //Set Alignment Settings for Style
            style2.HorizontalAlignment = TextAlignmentType.Right;

            //Set foreground color
            style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);

            //Set Pattern for Style
            style2.Pattern = BackgroundType.Solid;

            //Apply Style2 to cells A2 and A4
            cells["A2"].SetStyle(style2);
            cells["A4"].SetStyle(style2);


            //Initialize Style3
            Style style3 = workbook.Styles[workbook.Styles.Add()];

            //copy properties from Style2
            style3.Copy(style2);

            //Set cell format
            style3.Custom = "\"$\"#,##0";

            //Loop cells and Set Style3
            for (int i = 1; i <= 3; i++)
            {
                if (i % 2 != 0)
                {
                    cells[i, 1].SetStyle(style3);
                    cells[i, 2].SetStyle(style3);
                    cells[i, 3].SetStyle(style3);
                    cells[i, 4].SetStyle(style3);
                    cells[i, 5].SetStyle(style3);
                }
            }

            //Initialize Style4
            Style style4 = workbook.Styles[workbook.Styles.Add()];

            //Copy properties from style2
            style4.Copy(style2);

            //Sets foreground color
            style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);

            //Set Style Pattern
            style4.Pattern = BackgroundType.Solid;

            //Apply Style4 to Cell A3
            cells["A3"].SetStyle(style4);


            //Iniatalize Style5
            Style style5 = workbook.Styles[workbook.Styles.Add()];

            //Copy Style4 properties
            style5.Copy(style4);

            //Set cell format
            style5.Custom = "\"$\"#,##0";

            //Loop Cells ans set STyle
            for (int i = 1; i <= 3; i++)
            {
                if (i % 2 == 0)
                {
                    cells[i, 1].SetStyle(style5);
                    cells[i, 2].SetStyle(style5);
                    cells[i, 3].SetStyle(style5);
                    cells[i, 4].SetStyle(style5);
                    cells[i, 5].SetStyle(style5);
                }
            }
        }

		private void CreateStaticReport(Workbook workbook)
		{
            //Initialize Worksheet
			Worksheet sheet = workbook.Worksheets[0];

			//Set the name of worksheet
			sheet.Name = "Line";

            //Set Gridlines invisible
			sheet.IsGridlinesVisible = false;

			//Create chart depending on the Chart Type selected from List On UI
			int chartIndex = 0; 
			switch ( ChartTypeList.SelectedItem.Text )
			{
				case "Line":
					chartIndex = sheet.Charts.Add(ChartType.Line,5,1,29,10);
					break;				
				case "LineWithDataMarkers":
					chartIndex = sheet.Charts.Add(ChartType.LineWithDataMarkers,5,1,29,10);
					break;
			}

            //Initialize Chart Object
			Chart chart = sheet.Charts[chartIndex];	

            //Set Chart's MajorGridLines invisib;e
			chart.CategoryAxis.MajorGridLines.IsVisible = false;

			//Set Properties of chart title 
			chart.Title.Text = "Sales By Region For Years";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;

			//Set Properties of nseries
			chart.NSeries.Add("B2:F4", false);

            //Set Datasource for Nseries Category
			chart.NSeries.CategoryData = "B1:F1";

            //Set Nseries color varience to True
			chart.NSeries.IsColorVaried = true;	


            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

            //Loop over the cells
			for ( int i = 0 ;i < chart.NSeries.Count; i ++ )
			{	
                //Set Nseriese Name from Value in cells
				chart.NSeries[i].Name = cells[i+1,0].Value.ToString();

                //Selection from dropdownlist by name NSeriesMarkerStyle
                //Set MarkerStyle depending on selection
				switch (NSeriesMarkerStyle.SelectedItem.Text)
				{
					case "Automatic":
						chart.NSeries[i].MarkerStyle = ChartMarkerType.Automatic;
						break;
					case "Circle":
						chart.NSeries[i].MarkerStyle = ChartMarkerType.Circle;
						break;
					case "Dash":
						chart.NSeries[i].MarkerStyle = ChartMarkerType.Dash;
						break;
					case "Diamond":
						chart.NSeries[i].MarkerStyle = ChartMarkerType.Diamond;
						break;
					case "Dot":
						chart.NSeries[i].MarkerStyle = ChartMarkerType.Dot;
						break;
					case "None":
						chart.NSeries[i].MarkerStyle = ChartMarkerType.None;
						break;
					case "Square":
						chart.NSeries[i].MarkerStyle = ChartMarkerType.Square;
						break;
					case "SquarePlus":
						chart.NSeries[i].MarkerStyle = ChartMarkerType.SquarePlus;
						break;
					case "SquareStar":
						chart.NSeries[i].MarkerStyle = ChartMarkerType.SquareStar;
						break;
					case "SquareX":
						chart.NSeries[i].MarkerStyle = ChartMarkerType.SquareX;
						break;
					case "Triangle":
						chart.NSeries[i].MarkerStyle = ChartMarkerType.Triangle;
						break;
				}				


                //Set Nseriese Marker Background and ForeGround colors
				chart.NSeries[i].MarkerBackgroundColor = Color.FromName( NMarkBackColor.SelectedItem.Text );
				chart.NSeries[i].MarkerForegroundColor = Color.FromName( NMarkForeColor.SelectedItem.Text ) ;

                //Set NSeries Marker size 
				chart.NSeries[i].MarkerSize = int.Parse( NSeriesMarkSize.SelectedItem.Text );				

				//Set Properties of categoryaxis title 
				chart.CategoryAxis.Title.Text = "Year(2002-2006)";
				chart.CategoryAxis.Title.TextFont.Color = Color.Black;
				chart.CategoryAxis.Title.TextFont.IsBold = true;
				chart.CategoryAxis.Title.TextFont.Size = 10;

				//Set the legend position to Top
				chart.Legend.Position = LegendPositionType.Top;
			}
		}
		
	}
}
