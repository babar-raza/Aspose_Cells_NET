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
	/// Summary description for OpenHighLowClose.
	/// </summary>
	public class OpenHighLowClose : System.Web.UI.Page
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
            workbook.Save(HttpContext.Current.Response, "OpenHighLowClose." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
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

            //Initilize Cells
            Cells cells = workbook.Worksheets[0].Cells;

            //Put values for Column Header
            cells["A1"].PutValue("Company Name");
            cells["B1"].PutValue("Open");
            cells["C1"].PutValue("High");
            cells["D1"].PutValue("Low");
            cells["E1"].PutValue("Close");

            //Put values for Row 1
            cells["A2"].PutValue("Microsoft");
            cells["B2"].PutValue(21.00);
            cells["C2"].PutValue(27.20);
            cells["D2"].PutValue(23.49);
            cells["E2"].PutValue(25.45);				


            //Put values for Row 2
            cells["A3"].PutValue("Mutual Fund 1");
            cells["B3"].PutValue(28.52);
            cells["C3"].PutValue(25.03);
            cells["D3"].PutValue(19.55);
            cells["E3"].PutValue(23.05);

            //Put values for Row 3
            cells["A4"].PutValue("Mutual Fund 2");
            cells["B4"].PutValue(9.05);
            cells["C4"].PutValue(19.05);
            cells["D4"].PutValue(15.12);
            cells["E4"].PutValue(17.32);
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
            cells.SetColumnWidth(0, 15);

            //Set Style for Header
            cells["A1"].SetStyle(style1);
            cells["B1"].SetStyle(style1);
            cells["C1"].SetStyle(style1);
            cells["D1"].SetStyle(style1);

            //Initialize Style 2
            Style style2 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style2.Copy(style1);

            //Set Font to Normal
            style2.Font.IsBold = false;

            //Set Style Alignment
            style2.HorizontalAlignment = TextAlignmentType.Right;

            //Set foreground color
            style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);

            //Set Style Pattern
            style2.Pattern = BackgroundType.Solid;

            //Set Style
            cells["A2"].SetStyle(style2);
            cells["A4"].SetStyle(style2);

            //Initialize Style 3
            Style style3 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style3.Copy(style2);

            //Set cell format
            style3.Number = 2;


            //Loop over the cells and Set Style
            for (int i = 1; i <= 3; i++)
            {
                if (i % 2 != 0)
                {
                    cells[i, 1].SetStyle(style3);
                    cells[i, 2].SetStyle(style3);
                    cells[i, 3].SetStyle(style3);
                }
            }

            //Initialize Style 4
            Style style4 = workbook.Styles[workbook.Styles.Add()];

            //Copy the properties from another Style Object
            style4.Copy(style2);

            //Set foreground color
            style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);

            //Set Style Pattern
            style4.Pattern = BackgroundType.Solid;

            //Apply Style
            cells["A3"].SetStyle(style4);

            //Initialize Style 4
            Style style5 = workbook.Styles[workbook.Styles.Add()];

            //Copy the properties from another Style Object
            style5.Copy(style4);

            //Set cell format
            style5.Number = 2;

            //Loop over cells and Set Style
            for (int i = 1; i <= 3; i++)
            {
                if (i % 2 == 0)
                {
                    cells[i, 1].SetStyle(style5);
                    cells[i, 2].SetStyle(style5);
                    cells[i, 3].SetStyle(style5);
                }
            }
        }

		private void CreateStaticReport(Workbook workbook)
		{  
            //Get index of newly added Worksheet
			int sheetIndex = workbook.Worksheets.Add();

            //Initialize Worksheet for given index
			Worksheet sheet = workbook.Worksheets[sheetIndex];
			
            //Set the name of worksheet
			sheet.Name = "Chart";

            //Create chart of Type 	StockOpenHighLowClose
			int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose,1,1,25,10);

            //Initialize Chart
			Chart chart = sheet.Charts[chartIndex];

			//Set pproperties of nseries
			chart.NSeries.Add("Data!B2:E4",true);
			chart.NSeries.CategoryData = "Data!A2:A4";

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

            //loop over NSeries
			for ( int i = 0 ;i < chart.NSeries.Count ; i ++)
			{				
                //Set Name from values of cells
				chart.NSeries[i].Name = cells[0,i+1].Value.ToString();
			}

			//Set properties of chart title
			chart.Title.Text = " Stock chart";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;

			//Set Properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Scock Names";
			chart.CategoryAxis.Title.TextFont.Color = Color.Black;
			chart.CategoryAxis.Title.TextFont.Size = 10;
			chart.CategoryAxis.Title.TextFont.IsBold = true;

			//Set properties of valueaxis title
			chart.ValueAxis.Title.Text= "Stock Price";
			chart.ValueAxis.Title.TextFont.Color = Color.Black;
			chart.ValueAxis.Title.TextFont.IsBold = true;
			chart.ValueAxis.Title.TextFont.Size =10;
			chart.ValueAxis.Title.RotationAngle = 90;			
		}		
	}
}
