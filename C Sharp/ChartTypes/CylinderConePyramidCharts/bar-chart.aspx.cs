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
	/// Summary description for BarChart.
	/// </summary>
	public class BarChart : System.Web.UI.Page
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
            workbook.Save(HttpContext.Current.Response, "3DBarChart." +  ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));		
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
            //Initialize worksheet
			Worksheet sheet = workbook.Worksheets[0];

            //Hide gridlines of worksheet
			sheet.IsGridlinesVisible = false;

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

			//Put value into cells
			cells.SetColumnWidth(0,13.00);
			cells["A1"].PutValue("Region");
			cells["B1"].PutValue("Attendance");
			cells["A2"].PutValue("Providence");
			cells["B2"].PutValue(120);
			cells["A3"].PutValue("Philadelphia");
			cells["B3"].PutValue(150);
			cells["A4"].PutValue("Atlanta");
			cells["B4"].PutValue(180);
			cells["A5"].PutValue("Charleston");
			cells["B5"].PutValue(330);
			cells["A6"].PutValue("Detroit");
			cells["B6"].PutValue(380);	
		}

        private void CreateCellsFormatting(Workbook workbook)
        {
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
            style1.HorizontalAlignment = TextAlignmentType.Center;
            style1.VerticalAlignment = TextAlignmentType.Center;

            Cells cells = workbook.Worksheets[0].Cells;

            //Set the width of the specified column 
            cells.SetColumnWidth(1, 13.00);

            cells["A1"].SetStyle(style1);
            cells["B1"].SetStyle(style1);

            Style style2 = workbook.Styles[workbook.Styles.Add()];
            //Copy data from another style object
            style2.Copy(style1);
            style2.Font.IsBold = false;
            style2.HorizontalAlignment = TextAlignmentType.Right;
            //Set foreground color
            style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
            style2.Pattern = BackgroundType.Solid;

            for (int i = 1; i <= 11; i++)
            {
                if (i % 2 != 0)
                {
                    cells[i, 0].SetStyle(style2);
                    cells[i, 1].SetStyle(style2);
                }
            }

            Style style3 = workbook.Styles[workbook.Styles.Add()];
            style3.Copy(style2);
            //Sets foreground color
            style3.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
            style3.Pattern = BackgroundType.Solid;

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
            //get worksheet index after adding new worksheet
			int sheetIndex = workbook.Worksheets.Add(SheetType.Chart);

            //intialize worksheet on given index
			Worksheet sheet = workbook.Worksheets[sheetIndex];

			//Set the name of worksheet
			sheet.Name = "3DBar Chart";

            //Create chart depending on selected value from ChartTypeList
			int chartIndex = 0;			
			switch (ChartTypeList.SelectedItem.Text)
			{
				case "CylindericalBar":
					chartIndex = sheet.Charts.Add(ChartType.CylindricalBar, 0, 0, 0, 0);
					break;
				case "ConicalBar":
					chartIndex = sheet.Charts.Add(ChartType.ConicalBar,0,0,0,0);
					break;
				case "PyramidBar":
					chartIndex = sheet.Charts.Add(ChartType.PyramidBar,0,0,0,0);
					break;
			}	
			
            //Initialize Chart
			Chart chart = sheet.Charts[chartIndex];

			//Set properties of chart not to show Border
			chart.PlotArea.Border.IsVisible = false;

            //Set properties of chart not to show legend
			chart.ShowLegend = false;

			//Set properties of chart title
			chart.Title.Text = "Attendance By Region";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;

			//Set properties of nseries
			chart.NSeries.Add("Sheet1!B2:B6",true);			
			chart.NSeries.CategoryData = "Sheet1!A2:A6";			

			//Set properties of valueaxis title
			chart.ValueAxis.Title.Text = "Attendance";
			chart.ValueAxis.Title.TextFont.Color = Color.Black;
			chart.ValueAxis.Title.TextFont.IsBold = true;
			chart.ValueAxis.Title.TextFont.Size = 10;
			
			//Set properties of categoryaxis
			chart.CategoryAxis.IsPlotOrderReversed = true;			
		}

	}
}
