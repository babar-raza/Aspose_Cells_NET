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
	/// Summary description for PercentStackedLine.
	/// </summary>
	public class PercentStackedLine : System.Web.UI.Page
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
            workbook.Save(HttpContext.Current.Response, "PercentStackedLine." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

        private void CreateStaticData(Workbook workbook)
        {
            //Initialize Worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            //Set the name of worksheet
            worksheet.Name = "Data";

            //Set Gridlines invisible
            worksheet.IsGridlinesVisible = false;

            Cells cells = workbook.Worksheets[0].Cells;
            //Put values in row 1
            cells["A1"].PutValue("Product Name");
            cells["B1"].PutValue("Quarter1");
            cells["C1"].PutValue("Quarter2");
            cells["D1"].PutValue("Quarter3");
            cells["E1"].PutValue("Quarter4");

            //Put values in row 2
            cells["A2"].PutValue("Product1");
            cells["B2"].PutValue(0.33);
            cells["C2"].PutValue(0.21);
            cells["D2"].PutValue(0.35);
            cells["E2"].PutValue(0.22);

            //Put values in row 3
            cells["A3"].PutValue("Product2");
            cells["B3"].PutValue(0.17);
            cells["C3"].PutValue(0.54);
            cells["D3"].PutValue(0.17);
            cells["E3"].PutValue(0.60);

            //Put values in row 4
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

            //Set border settings for Style1
            style1.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style1.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style1.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style1.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            //Set Font IsBold property to True
            style1.Font.IsBold = true;

            //Set Style Alignment
            style1.HorizontalAlignment = TextAlignmentType.Center;
            style1.VerticalAlignment = TextAlignmentType.Center;

            //Initialize Cells
            Cells cells = workbook.Worksheets[0].Cells;

            //Set the width of the specified column 
            cells.SetColumnWidth(0, 15);

            //Apply style Top Row
            cells["A1"].SetStyle(style1);
            cells["B1"].SetStyle(style1);
            cells["C1"].SetStyle(style1);
            cells["D1"].SetStyle(style1);
            cells["E1"].SetStyle(style1);

            //Initialize Style2
            Style style2 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style2.Copy(style1);

            //Set Font Style
            style2.Font.IsBold = false;

            //Set Style Alignment
            style2.HorizontalAlignment = TextAlignmentType.Right;

            //Set foreground color
            style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);

            //Set Style Pattern
            style2.Pattern = BackgroundType.Solid;

            //Set style on Cells A2 and A4
            cells["A2"].SetStyle(style2);
            cells["A4"].SetStyle(style2);


            //initialize Style3
            Style style3 = workbook.Styles[workbook.Styles.Add()];

            //Copy Style Properties from another Style
            style3.Copy(style2);

            //Set cell format
            style3.Number = 9;


            //loop over the cells and set the style 
            for (int i = 1; i <= 3; i++)
            {
                if (i % 2 != 0)
                {
                    cells[i, 0].SetStyle(style2);
                    cells[i, 1].SetStyle(style3);
                    cells[i, 2].SetStyle(style3);
                    cells[i, 3].SetStyle(style3);
                    cells[i, 4].SetStyle(style3);
                }
            }


            //Initialize the Style4
            Style style4 = workbook.Styles[workbook.Styles.Add()];

            //Copy the style properties from another style
            style4.Copy(style2);

            //Sets foreground color
            style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);

            //Set Style Pattern
            style4.Pattern = BackgroundType.Solid;

            //Set Style on Cell A3
            cells["A3"].SetStyle(style4);

            //initialize Style5
            Style style5 = workbook.Styles[workbook.Styles.Add()];

            //Copy the style properties from another style
            style5.Copy(style4);

            //Set cell format
            style5.Number = 9;

            //loop over the cells and set the style
            for (int i = 1; i <= 3; i++)
            {
                if (i % 2 == 0)
                {
                    cells[i, 0].SetStyle(style5);
                    cells[i, 1].SetStyle(style5);
                    cells[i, 2].SetStyle(style5);
                    cells[i, 3].SetStyle(style5);
                    cells[i, 4].SetStyle(style5);
                }
            }
        }

		private void CreateStaticReport(Workbook workbook)
		{
            //get index of newly added worksheet
			int sheetIndex = workbook.Worksheets.Add(SheetType.Chart);

            //initialize worksheet for given index
			Worksheet sheet = workbook.Worksheets[sheetIndex];

			//Set the name of worksheet
			sheet.Name = "Chart";

            //Create Chart depending on selection made on ChartTypeList
			int chartIndex = 0;	
			switch (ChartTypeList.SelectedItem.Text)
			{
				case "Line100PercentStacked":
					chartIndex = sheet.Charts.Add(ChartType.Line100PercentStacked,0,0,0,0);
					break;
				case "Line100PercentStackedWithDataMarkers":
					chartIndex = sheet.Charts.Add(ChartType.Line100PercentStackedWithDataMarkers,0,0,0,0);
					break;
			}

            //Inialize Chart Object
			Chart chart = sheet.Charts[chartIndex];

			//Set properties of chart
			chart.CategoryAxis.MajorGridLines.IsVisible = false;

			//Set properies to chart title
			chart.Title.Text = "Product contribution to total sales";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;

			//Set properies to nseries
			chart.NSeries.Add("Data!B2:E4", false);

            //Set Nseriese Category Datasource
			chart.NSeries.CategoryData = "Data!B1:E1";

            //Initalize Cells
			Cells cells = workbook.Worksheets[0].Cells;

            //loop over the cells
			for ( int i = 0 ; i < chart.NSeries.Count ; i ++ )
			{
                //name NSeries from values in Cells
				chart.NSeries[i].Name = cells[i+1,0].Value.ToString();
			}

			//Set properies to valueaxis
			chart.ValueAxis.Title.Text = "% of total sales";
			chart.ValueAxis.Title.TextFont.Color = Color.Black;
			chart.ValueAxis.Title.TextFont.IsBold = true;
			chart.ValueAxis.Title.TextFont.Size = 10;
			chart.ValueAxis.Title.RotationAngle = 90;
		}

	}
}
