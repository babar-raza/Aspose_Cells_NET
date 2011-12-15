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
	/// Summary description for StackedBar.
	/// </summary>
	public class StackedBar : System.Web.UI.Page
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
            workbook.Save(HttpContext.Current.Response, "StackedBar." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));	
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
			Worksheet sheet = workbook.Worksheets[0];
			//Set the name of worksheet
			sheet.Name = "Data";
			sheet.IsGridlinesVisible = false;			

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

			//Put string values in row cells of column 1
			cells["A1"].PutValue("Region");
			cells["A2"].PutValue("France");
			cells["A3"].PutValue("Germany");
			cells["A4"].PutValue("English");
			cells["A5"].PutValue("Italy");

			//Put Number values in row cells of column 2
			cells["B2"].PutValue(25000);
			cells["B3"].PutValue(15000);
			cells["B4"].PutValue(30000);
			cells["B5"].PutValue(20000);

            //Put Number values in row cells of column 3
			cells["C2"].PutValue(20000);
			cells["C3"].PutValue(15000);
			cells["C4"].PutValue(25000);
			cells["C5"].PutValue(30000);

            //Put Number values in row cells of column 4
			cells["D2"].PutValue(30000);
			cells["D3"].PutValue(32000);
			cells["D4"].PutValue(15000);
			cells["D5"].PutValue(10000);

            //Put string values in cells B1, C1, D1
			cells["B1"].PutValue("Apple");
			cells["C1"].PutValue("Orange");
			cells["D1"].PutValue("Banana");
		}

		private void CreateCellsFormatting(Workbook workbook)
		{
            //Intialize Style1
			Style style1 = workbook.Styles[workbook.Styles.Add()];
			
            //Set border settings for Style1
			style1.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
			
            //Set Font property IsBold to True
            style1.Font.IsBold = true;

            //Set Style Alignment
			style1.HorizontalAlignment = TextAlignmentType.Center;			

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

            //Set Style for First Row
			cells["A1"].SetStyle(style1);
			cells["B1"].SetStyle(style1);
			cells["C1"].SetStyle(style1);
			cells["D1"].SetStyle(style1);
			
            //Initialize Style 2
			Style style2 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy data from another style object
			style2.Copy(style1);
			
            //Set Font IsBold Property to false
            style2.Font.IsBold = false;

            //Set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right;
			
            //Set foreground color
			style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
			
            //Set Style Pattern
            style2.Pattern = BackgroundType.Solid;	

            //Set Style of cell A2 and A4
			cells["A2"].SetStyle(style2);
			cells["A4"].SetStyle(style2);

            //Initialize Style3
			Style style3 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy properties from Style2
            style3.Copy(style2);

			//Set cell format
			style3.Custom = "\"$\"#,##0";	
		
			//Loop Over the cells and set Style
            for(int i = 1; i <= 4; i ++)
			{
				if ( i % 2 !=0)
				{
					cells[i,1].SetStyle(style3);
					cells[i,2].SetStyle(style3);
					cells[i,3].SetStyle(style3);
				}
			}	

            //Initialize Style4
			Style style4 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy the Properties of Style2
            style4.Copy(style2);

			//Set foreground color
			style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
			
            //Set Style Pattern
            style4.Pattern = BackgroundType.Solid;	

            //Apply Style to cells A3 and A5
			cells["A3"].SetStyle(style4);
			cells["A5"].SetStyle(style4);

            //Initialize Style5
			Style style5 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy the properties of Style4
            style5.Copy(style4);

			//Set cell format
			style5.Custom = "\"$\"#,##0";	

            //Loop over the cells as set the Style
			for ( int i = 1; i <= 4; i ++ )
			{
				if (i % 2 == 0)
				{	
					cells[i,1].SetStyle(style5);
					cells[i,2].SetStyle(style5);
					cells[i,3].SetStyle(style5);
				}
			}	
		}
		
		private void CreateStaticReport(Workbook workbook)
		{
            //get the next index for worksheets in workbook
			int sheetIndex = workbook.Worksheets.Add();

            //Initialize worksheet from given index
			Worksheet sheet = workbook.Worksheets[sheetIndex];
			
            //Set the name of worksheet
			sheet.Name = "Chart";				

			//Create chart
			int indexChart = 0;
            //Create chart, If Check box on Ui is Checked then create Bar3DStacked Chart else BarStacked Chart
            if (CheckBoxShow3D.Checked)
				indexChart = sheet.Charts.Add (ChartType.Bar3DStacked,1,1,21,10);	
			else
				indexChart = sheet.Charts.Add (ChartType.BarStacked,1,1,21,10);	
			Chart chart = sheet.Charts[indexChart];

            //Set properties of chart and hide gridLines based upon the state od check box on UI
			if (CheckBoxShow3D.Checked)
                chart.PlotArea.Border.IsVisible = false;
			chart.CategoryAxis.MajorGridLines.IsVisible = false;

			//Set properties of chart title
			chart.Title.Text = "Fruit Sales By Region";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;

			//Set properties of nseries
			chart.NSeries.Add ("Data!B2:D5", true);
			chart.NSeries.CategoryData = "Data!A2:A5";
			chart.NSeries.IsColorVaried = true;

            //Initalize Cells
			Cells cells = workbook.Worksheets[0].Cells;

            //loop over the Chart's Nseries and Assign Name from Cell Values
			for ( int i = 0; i< chart.NSeries.Count; i ++ )
			{
				chart.NSeries[i].Name = cells[0,i+1].Value.ToString(); 
			}

			//Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Region";
			chart.CategoryAxis.Title.TextFont.Color = Color.Black;
			chart.CategoryAxis.Title.TextFont.IsBold = true;
			chart.CategoryAxis.Title.TextFont.Size = 10;
			chart.CategoryAxis.Title.RotationAngle = 90;			

			//Set properties of legend
			chart.Legend.Position = LegendPositionType.Top;
		}       	
	}
}
