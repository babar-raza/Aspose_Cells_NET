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


namespace Aspose.Cells.Demos.ChartTypes._3DCharts
{
	/// <summary>
	/// Summary description for Column3D.
	/// </summary>
	public class Column3D : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.Button btnProcess;
		protected System.Web.UI.WebControls.DropDownList ChartTypeList;
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
            workbook.Save(HttpContext.Current.Response, "Column3D." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
			Worksheet sheet = workbook.Worksheets[0];
			sheet.IsGridlinesVisible = false;

			Cells cells = workbook.Worksheets[0].Cells;

			//Put values in row cells of Column 1
			cells["A1"].PutValue("Region");
			cells["B1"].PutValue("Apple");
			cells["C1"].PutValue("Orange");

            //Put values in row cells of Column 2
			cells["A2"].PutValue("France");
			cells["B2"].PutValue(800000);
			cells["C2"].PutValue(300000);

            //Put values in row cells of Column 3
			cells["A3"].PutValue("Germany");
			cells["B3"].PutValue(200000);
			cells["C3"].PutValue(600000);

            //Put values in row cells of Column 4
			cells["A4"].PutValue("England");		
			cells["B4"].PutValue(400000);
			cells["C4"].PutValue(600000);			
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

            //Set Style of Column Headers
			cells["A1"].SetStyle(style1);
			cells["B1"].SetStyle(style1);
			cells["C1"].SetStyle(style1);

            //Initialize Style 2
			Style style2 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style2.Copy(style1);
			
            //Set FOnt IsBold property to False
            style2.Font.IsBold = false;

            //Set Style Alignmment
			style2.HorizontalAlignment = TextAlignmentType.Right;
			
            //Set foreground color
			style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
			
            //Set Style Patern
            style2.Pattern = BackgroundType.Solid;

            //Apply Style on Cells A2 and A4
			cells["A2"].SetStyle(style2);
			cells["A4"].SetStyle(style2);

            //Initialize Style 3
			Style style3 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style3.Copy(style2);
			
            //Set cell format
			style3.Custom = "\"$\"#,##0";

            //Apply Style on Cells B2, C2, B4 and C4
			cells["B2"].SetStyle(style3);
			cells["C2"].SetStyle(style3);
			cells["B4"].SetStyle(style3);
			cells["C4"].SetStyle(style3);

            //Initialize Style 4
			Style style4 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style4.Copy(style3);
			
            //Set foreground color
			style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
			
            style4.Pattern = BackgroundType.Solid;

            //Apply Style on Cells A3
            cells["A3"].SetStyle(style4);

            //Initialize Style 5
			Style style5 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
			style5.Copy(style4);
			
            //Set cell format
			style5.Custom = "\"$\"#,##0";

            //Set Style of cells B3 and C3
			cells["B3"].SetStyle(style5);
			cells["C3"].SetStyle(style5);			
		}

		private void CreateStaticReport(Workbook workbook)
		{
            //Get index of newly added worksheet
			int sheetIndex = workbook.Worksheets.Add(SheetType.Chart);

            //initialize Worksheet
			Worksheet sheet = workbook.Worksheets[sheetIndex];

			//Set the name of worksheet
			sheet.Name = "Column3D Chart";

            //Create Chart depending on selected value on ChartTypeList
			int indexChart = 0; 
			switch (ChartTypeList.SelectedItem.Text)
			{
				case "Cylinder":
					indexChart = sheet.Charts.Add(ChartType.CylindricalColumn3D, 0, 0, 0, 0);	
					break;
				case "Cone":
					indexChart = sheet.Charts.Add(ChartType.ConicalColumn3D,0,0,0,0);
					break;
				case "Pyramid":
					indexChart = sheet.Charts.Add(ChartType.PyramidColumn3D,0,0,0,0);
					break;
			}			

			Chart chart = sheet.Charts[indexChart];
			chart.PlotArea.Border.IsVisible = false;

			//Set Properties of chart title
			chart.Title.Text = "Fruit Sales By Region";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;

			//Set properties of nseries
			chart.NSeries.Add("Sheet1!B2:C4",true);

            //Set nseries Category Data source
			chart.NSeries.CategoryData = "Sheet1!A2:A4";

            //Set nseries Color varience to True
			chart.NSeries.IsColorVaried = true;
			
            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

            //Loop on Nseriese and Name them as values in cells
			for ( int i = 0 ; i < chart.NSeries.Count ; i ++)
			{
				chart.NSeries[i].Name = cells[0,i+1].Value.ToString();
			}

			//Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Region";
			chart.CategoryAxis.Title.TextFont.Color = Color.Black;
			chart.CategoryAxis.Title.TextFont.IsBold = true;
			chart.CategoryAxis.Title.TextFont.Size = 10;			

			//Set properties of legend
			chart.Legend.Position = LegendPositionType.Top;
		}
	}
}
