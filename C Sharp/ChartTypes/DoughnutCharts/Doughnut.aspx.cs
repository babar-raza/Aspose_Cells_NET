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
	/// Summary description for Doughnut.
	/// </summary>
	public class Doughnut : System.Web.UI.Page
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
			this.ID = "Doughnut";
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
            workbook.Save(HttpContext.Current.Response, "Doughnut." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
            //initialize cells of worksheet
			Cells cells = workbook.Worksheets[0].Cells;

			//Put values in row cells of column 1
			cells["A1"].PutValue("Product Name");
			cells["A2"].PutValue("Apple");
			cells["A3"].PutValue("Orange");

            //Put values in some cells
			cells["B1"].PutValue(2005);
			cells["C1"].PutValue(2006);				

			cells["B2"].PutValue(30000);
			cells["C2"].PutValue(50000);

			cells["B3"].PutValue(10000);
			cells["C3"].PutValue(20000);
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
			style1.VerticalAlignment = TextAlignmentType.Center;

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

			//Set the width of the specified column 
			cells.SetColumnWidth(0,15);

            //Apply Style to Column Header (first row)
			cells["A1"].SetStyle(style1);
			cells["B1"].SetStyle(style1);
			cells["C1"].SetStyle(style1);			

            //Initialize Style2
			Style style2 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy data from another style object
			style2.Copy(style1);
			
            //Set IsBold property of Font to False
            style2.Font.IsBold = false;

            //Set Alignment of Style
			style2.HorizontalAlignment = TextAlignmentType.Right;
			
            //Set foreground color
			style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
			style2.Pattern = BackgroundType.Solid;

            //Apply style to A2
			cells["A2"].SetStyle(style2);

            //Initialize Style2
			Style style3 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style3.Copy(style2);

			//Set cell format
			style3.Custom = "\"$\"#,##0";

            //Apply style to B2, C2
			cells["B2"].SetStyle(style3);
			cells["C2"].SetStyle(style3);

            //Initialize Style2
			Style style4 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style4.Copy(style2);
			
            //Set foreground color
			style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
			
            style4.Pattern = BackgroundType.Solid;

            //Apply style to A3
			cells["A3"].SetStyle(style4);

            //Initialize Style2
			Style style5 = workbook.Styles[workbook.Styles.Add()];

            //Copy data from another style object
            style5.Copy(style4);

			//Set cell format
			style5.Custom = "\"$\"#,##0";	

            //Apply style to B3
			cells["B3"].SetStyle(style5);
			cells["C3"].SetStyle(style5);
		}

		private void CreateStaticReport(Workbook workbook)
		{		
	        //Initialize Worksheet
			Worksheet sheet = workbook.Worksheets[0];
			
            //Set the name of worksheet
			sheet.Name = "Doughnut";
			
            //Set Gridlines to Invisible for Worksheet [0]
            sheet.IsGridlinesVisible = false;

            //Create chart of type Doughnut
			int chartIndex = sheet.Charts.Add(ChartType.Doughnut,3,1,27,10);		
			Chart chart = sheet.Charts[chartIndex];

			//Set the properties of chart
			chart.PlotArea.Border.IsVisible = false;

			//Set the properties of chart title
			chart.Title.Text = "Fruit Sales by Region For Years";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;			

			//Set the properties of nseries
			chart.NSeries.Add("B2:C3",true);
			chart.NSeries.CategoryData = "B1:C1";
			chart.NSeries.IsColorVaried = true;			
			
            //loop over Nseries and Set DataLabels to show
			for ( int i = 0; i < chart.NSeries.Count ; i++)
			{
				chart.NSeries[i].DataLabels.ShowValue = true;
			}			

			//Set the properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Region";
			chart.CategoryAxis.Title.TextFont.Color = Color.Black;
			chart.CategoryAxis.Title.TextFont.IsBold = true;
			chart.CategoryAxis.Title.TextFont.Size = 10;

			//Set the properties of valueaxis title
			chart.ValueAxis.Title.Text = "Thousand";
			chart.ValueAxis.Title.TextFont.Color = Color.Black;
			chart.ValueAxis.Title.TextFont.IsBold = true;
			chart.ValueAxis.Title.TextFont.Size = 10;

			//Set the properties of legend to show on top
			chart.Legend.Position = LegendPositionType.Top;
		}
	}
}
