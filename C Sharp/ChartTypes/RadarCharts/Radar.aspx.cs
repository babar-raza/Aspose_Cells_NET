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
	/// Summary description for Radar.
	/// </summary>
	public class Radar : System.Web.UI.Page
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
			this.ID = "Radar";
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
            workbook.Save(HttpContext.Current.Response, "Radar." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));	
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

        private void CreateStaticData(Workbook workbook)
        {
            Worksheet sheet = workbook.Worksheets[0];
            //Set the name of worksheet
            sheet.Name = "Data";
            sheet.IsGridlinesVisible = false;

            Cells cells = workbook.Worksheets[0].Cells;
            //Put values fo row 1 (column Header)
            cells["A1"].PutValue("Brand Name");
            cells["B1"].PutValue("Vitamin A");
            cells["C1"].PutValue("Vitamin B1");
            cells["D1"].PutValue("Vitamin B2");
            cells["E1"].PutValue("Vitamin C");
            cells["F1"].PutValue("Vitamin D");
            cells["G1"].PutValue("Vitamin E");

            //Put Values for row 2
            cells["A2"].PutValue("Brand A");
            cells["B2"].PutValue(100);
            cells["C2"].PutValue(100);
            cells["D2"].PutValue(100);
            cells["E2"].PutValue(80);
            cells["F2"].PutValue(100);
            cells["G2"].PutValue(70);

            //put values for row 3
            cells["A3"].PutValue("Brand B");
            cells["B3"].PutValue(80);
            cells["C3"].PutValue(75);
            cells["D3"].PutValue(80);
            cells["E3"].PutValue(100);
            cells["F3"].PutValue(50);
            cells["G3"].PutValue(15);

            //put values for row 4
            cells["A4"].PutValue("Brand C");
            cells["B4"].PutValue(40);
            cells["C4"].PutValue(25);
            cells["D4"].PutValue(40);
            cells["E4"].PutValue(55);
            cells["F4"].PutValue(30);
            cells["G4"].PutValue(10);
        }

        private void CreateCellsFormatting(Workbook workbook)
        {
            //Initialize Style1
            Style style1 = workbook.Styles[workbook.Styles.Add()];
            //Set border for Style
            style1.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style1.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style1.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style1.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
            style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            //Set Style Font to Bold
            style1.Font.IsBold = true;

            //Set Style Alignment
            style1.HorizontalAlignment = TextAlignmentType.Center;


            //Initialize Cells
            Cells cells = workbook.Worksheets[0].Cells;

            //Set column Width
            cells.SetColumnWidth(0, 12);

            //loop over the cells 
            for (int i = 1; i <= 6; i++)
            {
                //Set the width of the specified column 
                cells.SetColumnWidth(i, 9);
            }

            //Set style of column header
            cells["A1"].SetStyle(style1);
            cells["B1"].SetStyle(style1);
            cells["C1"].SetStyle(style1);
            cells["D1"].SetStyle(style1);
            cells["E1"].SetStyle(style1);
            cells["F1"].SetStyle(style1);
            cells["G1"].SetStyle(style1);

            //initialize Style 2
            Style style2 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style2.Copy(style1);

            //set style font to Normal
            style2.Font.IsBold = false;

            //Set foreground color
            style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);

            //Set Style pattern
            style2.Pattern = BackgroundType.Solid;

            //Set Style Alignment
            style2.HorizontalAlignment = TextAlignmentType.Right;

            //loop over the cells and Set style
            for (int i = 1; i <= 3; i++)
            {
                if (i % 2 != 0)
                {
                    cells[i, 0].SetStyle(style2);
                    cells[i, 1].SetStyle(style2);
                    cells[i, 2].SetStyle(style2);
                    cells[i, 3].SetStyle(style2);
                    cells[i, 4].SetStyle(style2);
                    cells[i, 5].SetStyle(style2);
                    cells[i, 6].SetStyle(style2);
                }
            }

            //initialize Style 
            Style style3 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style3.Copy(style2);

            //Set Style ForegroundColor
            style3.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);

            //Set Style Pattern
            style3.Pattern = BackgroundType.Solid;

            //Loop over the cells and Set Style
            for (int i = 1; i <= 3; i++)
            {
                if (i % 2 == 0)
                {
                    cells[i, 0].SetStyle(style2);
                    cells[i, 1].SetStyle(style3);
                    cells[i, 2].SetStyle(style3);
                    cells[i, 3].SetStyle(style3);
                    cells[i, 4].SetStyle(style3);
                    cells[i, 5].SetStyle(style3);
                    cells[i, 6].SetStyle(style3);
                }
            }
        }

		private void CreateStaticReport(Workbook workbook)
        {
            //get index of newly added Worksheet
            int sheetIndex = workbook.Worksheets.Add();

            //initialize worksheet for given Index
            Worksheet sheet = workbook.Worksheets[sheetIndex];

            //Set the name of worksheet
            sheet.Name = "Chart";

            //Create chart depending on the ChartTypeList's SelectedItem
			int chartIndex = 0;
			switch (ChartTypeList.SelectedItem.Text)
			{
				case "Radar":
					chartIndex = sheet.Charts.Add(ChartType.Radar,5,1,29,10);
					break;				
				case "RadarWithDataMarkers":
					chartIndex = sheet.Charts.Add(ChartType.RadarWithDataMarkers,5,1,29,10);
					break;
			}				
			Chart chart = sheet.Charts[chartIndex];

			//Set properties of chart
			chart.PlotArea.Border.IsVisible = false;			

			//Set properties of chart title
			chart.Title.Text = "Nutritional Analysis";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;

			//Set properties of nseries
			chart.NSeries.Add("B2:G4",false);
			chart.NSeries.CategoryData = "B1:G1";

            //Initialize Cells
            Cells cells = workbook.Worksheets[0].Cells;

            //loop over the NSeries
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                //Set NSeries Name to values from cells
                chart.NSeries[i].Name = cells["A" + (i + 2).ToString()].Value.ToString();
            }		
		}
		
	}
}
