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
	/// Summary description for Scatter.
	/// </summary>
	public class Scatter : System.Web.UI.Page
	{
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
			this.ID = "Scatter";
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
            workbook.Save(HttpContext.Current.Response, "Scatter." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));	
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

        private void CreateStaticData(Workbook workbook)
        {
            //Initialize Worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //name Worksheet
            sheet.Name = "Data";

            //Set Worksheet's gridlines invisible
            sheet.IsGridlinesVisible = false;

            //Initialize Cells
            Cells cells = workbook.Worksheets[0].Cells;

            //Put string in cells to make Column Header
            cells["A1"].PutValue("Daily Rainfall");
            cells["B1"].PutValue("Particulate");

            //Put values to make rows
            cells["A2"].PutValue(1.9);
            cells["B2"].PutValue(137);
            cells["A3"].PutValue(3.6);
            cells["B3"].PutValue(128);
            cells["A4"].PutValue(4.1);
            cells["B4"].PutValue(122);
            cells["A5"].PutValue(4.3);
            cells["B5"].PutValue(117);
            cells["A6"].PutValue(5);
            cells["B6"].PutValue(114);
            cells["A7"].PutValue(5.4);
            cells["B7"].PutValue(114);
            cells["A8"].PutValue(5.7);
            cells["B8"].PutValue(112);
            cells["A9"].PutValue(5.9);
            cells["B9"].PutValue(110);
            cells["A10"].PutValue(7.3);
            cells["B10"].PutValue(104);
        }

        private void CreateCellsFormatting(Workbook workbook)
        {
            //Initialize Style1
            Style style1 = workbook.Styles[workbook.Styles.Add()];

            //Set border style
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

            //Set Style Alignment
            style1.HorizontalAlignment = TextAlignmentType.Center;

            //Initialize Cells
            Cells cells = workbook.Worksheets[0].Cells;

            //Set the width of the specified column 
            cells.SetColumnWidth(0, 12);
            cells.SetColumnWidth(1, 10);

            //Set Style for Column Header 
            cells["A1"].SetStyle(style1);
            cells["B1"].SetStyle(style1);

            //Initialize Style2
            Style style2 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style2.Copy(style1);

            //Set Font to Not Bold
            style2.Font.IsBold = false;

            //Set foreground color
            style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);

            //Set Style Patern
            style2.Pattern = BackgroundType.Solid;

            //Set Style Alignment
            style2.HorizontalAlignment = TextAlignmentType.Right;

            //loop over the cells and set Style
            for (int i = 1; i <= 9; i++)
            {
                if (i % 2 != 0)
                {
                    cells[i, 0].SetStyle(style2);
                    cells[i, 1].SetStyle(style2);
                }
            }

            //Initialize Style3
            Style style3 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style3.Copy(style2);

            //Set foreground color
            style3.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);

            //Set Style pattern
            style3.Pattern = BackgroundType.Solid;

            //loop over the cells and set Style
            for (int i = 1; i <= 9; i++)
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
            //Initialize Worksheet
			Worksheet sheet = workbook.Worksheets[0];
			
            //Set the name of worksheet
			sheet.Name = "Scatter";
			
            //Set Worksheet's GridLines invisible
            sheet.IsGridlinesVisible = false;

			//Create chart of Type Scatter
			int chartIndex = sheet.Charts.Add(ChartType.Scatter,1,3,25,12);					
			
            //Initialize Chart
            Chart chart = sheet.Charts[chartIndex];

			//Set properties of chart
			chart.CategoryAxis.MajorGridLines.IsVisible = false;

            //Set Legend Position of chart
			chart.Legend.Position = LegendPositionType.Top;

			//Set properties of chart title
			chart.Title.Text = "Scatter Chart:Particulate Levels in Rainfall";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;
			
			//Set properties of nseries
			chart.NSeries.Add ("B2:B10",true);
			chart.NSeries[0].XValues = "A2:A10";

            //Loop over the NSeries and set Name
			for ( int i = 0 ; i < chart.NSeries.Count ; i ++)
			{
				chart.NSeries[i].Name = "Particulate";
			}

            Cells cells = workbook.Worksheets[0].Cells;
			//Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = cells["A1"].Value.ToString();
			chart.CategoryAxis.Title.TextFont.Color = Color.Black;
			chart.CategoryAxis.Title.TextFont.IsBold = true;
			chart.CategoryAxis.Title.TextFont.Size = 10;

			//Set properties of valueaxis title
			chart.ValueAxis.Title.Text = cells["B1"].Value.ToString();
			chart.ValueAxis.Title.TextFont.Color = Color.Black;
			chart.ValueAxis.Title.TextFont.IsBold = true;
			chart.ValueAxis.Title.TextFont.Size = 10;
			chart.ValueAxis.Title.RotationAngle = 90;		
		}
		
	}
}
