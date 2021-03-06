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
	/// Summary description for PyramidBar.
	/// </summary>
	public class PyramidBar : System.Web.UI.Page
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
            workbook.Save(HttpContext.Current.Response, "PyramidBar." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
            //Initialize Worksheet
			Worksheet sheet = workbook.Worksheets[0];

			//Set the name of worksheet
			sheet.Name = "Data";
			
            //Set Gridlines invisible
            sheet.IsGridlinesVisible = false;

            //Initialize Cells of Worksheet[0]
            Cells cells = workbook.Worksheets[0].Cells;

            //Put a value into a cell
            cells["A1"].PutValue("Year");

            //Get Style Object 
            Aspose.Cells.Style style = cells["A1"].GetStyle();

            //Set Font IsBold property to True
            style.Font.IsBold = true;

            //Set Style for Column Header
            cells["A1"].SetStyle(style);

            //Put values in cells in rows of column A
            cells["A2"].PutValue(1996);
            cells["A3"].PutValue(1997);
            cells["A4"].PutValue(1998);
            cells["A5"].PutValue(1999);
            cells["A6"].PutValue(2000);
            cells["A7"].PutValue(2001);
            cells["A8"].PutValue(2002);
            cells["A9"].PutValue(2003);
            cells["A10"].PutValue(2004);
            cells["A11"].PutValue(2005);
            cells["A12"].PutValue(2006);

            //Put Values for Column Header B
            cells["B1"].PutValue("No.Employees");

            //Put values in cells in rows of column B
            cells["B2"].PutValue(4);
            cells["B3"].PutValue(6);
            cells["B4"].PutValue(8);
            cells["B5"].PutValue(9);
            cells["B6"].PutValue(15);
            cells["B7"].PutValue(25);
            cells["B8"].PutValue(31);
            cells["B9"].PutValue(48);
            cells["B10"].PutValue(55);
            cells["B11"].PutValue(98);
            cells["B12"].PutValue(113);
        
		}

        private void CreateCellsFormatting(Workbook workbook)
        {
            //initialize Style1
            Style style1 = workbook.Styles[workbook.Styles.Add()];
            //Set border Style
            style1.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style1.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style1.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style1.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            //Set Font IsBold Property to True for Header
            style1.Font.IsBold = true;

            //Set Style Alignment
            style1.HorizontalAlignment = TextAlignmentType.Center;
            style1.VerticalAlignment = TextAlignmentType.Center;

            //initialize Cells of Worksheet[0]
            Cells cells = workbook.Worksheets[0].Cells;

            //Set the width of the specified column 
            cells.SetColumnWidth(1, 13.00);

            //Set Style of Header
            cells["A1"].SetStyle(style1);
            cells["B1"].SetStyle(style1);

            //initialize Style2
            Style style2 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style2.Copy(style1);

            //Set Font IsBold property to false
            style2.Font.IsBold = false;

            //Set Style Alignment
            style2.HorizontalAlignment = TextAlignmentType.Right;

            //Set foreground color
            style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);

            //Set Style Pattern
            style2.Pattern = BackgroundType.Solid;


            //Loop over the cells and Set Style
            for (int i = 1; i <= 11; i++)
            {
                if (i % 2 != 0)
                {
                    cells[i, 0].SetStyle(style2);
                    cells[i, 1].SetStyle(style2);
                }
            }

            //initialize Style3
            Style style3 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style3.Copy(style2);

            //Sets foreground color
            style3.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);

            //Set Style Pattern
            style3.Pattern = BackgroundType.Solid;

            //Loop over the cells and set Style
            for (int i = 1; i <= 11; i++)
            {
                if (i % 2 == 0)
                {
                    cells[i, 0].SetStyle(style2);
                    cells[i, 1].SetStyle(style3);
                }
            }
        }

		private void CreateStaticReport(Workbook workbook)
		{
            //Get index of newly added Worksheet
			int sheetIndex = workbook.Worksheets.Add();

            //Initialize Worksheet from given index
			Worksheet sheet = workbook.Worksheets[sheetIndex];

			//Set the name of worksheet
			sheet.Name = "Chart";

            //Create chart depending on selection of ChartTypeList
			int chartIndex = 0;
			switch ( ChartTypeList.SelectedItem.Text)
			{
				case "PyramidBar": 			
					chartIndex = sheet.Charts.Add(ChartType.PyramidBar,1,1,25,10);					
					break;
				case "PyramidBarStacked":				
					chartIndex = sheet.Charts.Add(ChartType.PyramidBarStacked,1,1,25,10);			
					break;			
			}			
	
            //Initalize Chart
			Chart chart = sheet.Charts[chartIndex];	

			//Set properties of chart
			chart.Walls.ForegroundColor = Color.White;
			chart.Floor.ForegroundColor = Color.White;			
			chart.Elevation = 15;
			chart.RotationAngle = 20;
			chart.PlotArea.Border.IsVisible = false;
			chart.ShowLegend = false;			
			chart.GapWidth = 10;
			chart.DepthPercent = 280;	

			//Set properties of nseries 
			chart.NSeries.Add("Data!B2:C12",true);

            //Assign Category datasource
			chart.NSeries.CategoryData = "Data!A2:A12";			

			//Set properties of chart title 
			chart.Title.Text = "Number of Employees";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 11;

			//Set properties of categoryaxis title
			chart.CategoryAxis.Title.TextFont.Color = Color.Black;
			chart.CategoryAxis.Title.TextFont.Size = 10;
			chart.CategoryAxis.Title.TextFont.Name = "Arial";
			chart.CategoryAxis.MajorTickMark = TickMarkType.Outside;
			chart.CategoryAxis.MinorTickMark = TickMarkType.None;
			chart.CategoryAxis.TickLabelPosition = TickLabelPositionType.Low;
			chart.CategoryAxis.TickLabelSpacing = 2;
			chart.CategoryAxis.TickMarkSpacing = 1;				

			//Set properties of valueaxis 
			chart.ValueAxis.Title.TextFont.Color = Color.Black;
			chart.ValueAxis.Title.TextFont.Size = 10;
			chart.ValueAxis.Title.TextFont.Name = "Arial";
			chart.ValueAxis.MajorTickMark = TickMarkType.Outside;
			chart.ValueAxis.MinorTickMark = TickMarkType.None;
			chart.ValueAxis.TickLabelPosition = TickLabelPositionType.NextToAxis; 
		}
	}
}
