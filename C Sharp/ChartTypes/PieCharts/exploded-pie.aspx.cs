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
	/// Summary description for ExplodedPie.
	/// </summary>
	public class ExplodedPie : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.Button btnProcess;
		protected System.Web.UI.WebControls.CheckBox CheckShow3D;
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
            workbook.Save(HttpContext.Current.Response, "ExplodedPie." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
            //Initialize Worksheet
			Worksheet sheet = workbook.Worksheets[0];
			
            //Set the name of worksheet
			sheet.Name = "Data";
			
            //Set GridLines invisible
            sheet.IsGridlinesVisible = false;

            //Initialize Cells
            Cells cells = workbook.Worksheets[0].Cells;

            //Put values for row cells of Column 1
            cells["A1"].PutValue("Region");
            cells["A2"].PutValue("France");
            cells["A3"].PutValue("Germany");
            cells["A4"].PutValue("England");
            cells["A5"].PutValue("Sweden");
            cells["A6"].PutValue("Italy");
            cells["A7"].PutValue("Spain");
            cells["A8"].PutValue("Portugal");

            //Put values for row cells of Column 2
            cells["B1"].PutValue("Sale");
            cells["B2"].PutValue(70000);
            cells["B3"].PutValue(55000);
            cells["B4"].PutValue(30000);
            cells["B5"].PutValue(40000);
            cells["B6"].PutValue(35000);
            cells["B7"].PutValue(32000);
            cells["B8"].PutValue(10000);
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

            //Set Font IsBold property
            style1.Font.IsBold = true;

            //Set Style Alignment
            style1.HorizontalAlignment = TextAlignmentType.Center;


            //Initialize Cells
            Cells cells = workbook.Worksheets[0].Cells;

            //Set Style for A1 and B1
            cells["A1"].SetStyle(style1);
            cells["B1"].SetStyle(style1);

            //Initialize Style2
            Style style2 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style2.Copy(style1);

            //Set Font to Bold
            style2.Font.IsBold = false;

            //Set foreground color
            style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);

            //Set Style Pattern
            style2.Pattern = BackgroundType.Solid;

            //Set Style Alignment
            style2.HorizontalAlignment = TextAlignmentType.Right;

            //Loop over the cells
            for (int i = 1; i <= 7; i++)
            {
                if (i % 2 != 0)
                {
                    //Apply Style
                    cells[i, 0].SetStyle(style2);
                }
            }

            //Initialize Style
            Style style3 = workbook.Styles[workbook.Styles.Add()];

            //Copy the Style from another
            style3.Copy(style2);

            //Set cell format
            style3.Custom = "\"$\"#,##0";


            //loop over the cells and Set Style
            for (int i = 1; i <= 7; i++)
            {
                if (i % 2 != 0)
                {
                    cells[i, 1].SetStyle(style3);
                }
            }

            //Initialize Style4
            Style style4 = workbook.Styles[workbook.Styles.Add()];

            //Copy Style from another
            style4.Copy(style2);

            //Set foreground color
            style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);

            //Set Style pattern
            style4.Pattern = BackgroundType.Solid;


            //Loop over the cells and set Style
            for (int i = 1; i <= 7; i++)
            {
                if (i % 2 == 0)
                {
                    cells[i, 0].SetStyle(style4);
                }
            }

            //Initialize Style
            Style style5 = workbook.Styles[workbook.Styles.Add()];

            //Copy the Style from Another
            style5.Copy(style4);

            //Set cell format
            style5.Custom = "\"$\"#,##0";

            //Loop over the cells and set Style
            for (int i = 1; i <= 7; i++)
            {
                if (i % 2 == 0)
                {
                    cells[i, 1].SetStyle(style5);
                }
            }
        }		

		private void CreateStaticReport(Workbook workbook)
		{ 
            //get index of newly added Worksheet
			int sheetIndex = workbook.Worksheets.Add();

            //Initialize Worksheet on given index
			Worksheet sheet = workbook.Worksheets[sheetIndex];
			
            //Set the name of worksheet
			sheet.Name = "Chart";

            //Create chart depending on selection made on Pie3DExploded
			int chartIndex = 0;
			if (CheckShow3D.Checked)
			    chartIndex = sheet.Charts.Add(ChartType.Pie3DExploded,1,1,25,10);
			else
				chartIndex = sheet.Charts.Add(ChartType.PieExploded,1,1,25,10);		   
			
            //Initialize Chart
            Chart chart = sheet.Charts[chartIndex];

            //Set properties of chart like ForegroundColor
			chart.PlotArea.Area.ForegroundColor = Color.Coral;

            //Set properties of chart like Boder Visibility
			chart.PlotArea.Border.IsVisible = false;

			//Set properties of chart title
			chart.Title.Text = "Sales By Region";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;
			
			//Set properties of nseries
			chart.NSeries.Add("Data!B2:B8", true);
			chart.NSeries.CategoryData = "Data!A2:A8";
			chart.NSeries.IsColorVaried = true;				

			//Set the legend position to Top
			chart.Legend.Position = LegendPositionType.Right;
		}
	}
}
