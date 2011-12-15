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
	/// Summary description for ClusteredBar.
	/// </summary>
	public class ClusteredBar : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.CheckBox checkBoxShow3D;
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

        /// <summary>
        /// Initialize Workbook, insert dummy data
        /// Create chart based on dummy data
        /// save the file in format selected from UI
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
            workbook.Save(HttpContext.Current.Response, "ClusteredBar." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticData(Workbook workbook)
		{
            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;
			
            //Put values in rows of cloumn A
			cells["A1"].PutValue("Region");
			cells["A2"].PutValue("France");
			cells["A3"].PutValue("Germany");
			cells["A4"].PutValue("England");

            //Put values into a cell B1 and C1
			cells["B1"].PutValue("Apple");
			cells["C1"].PutValue("Orange");

			//Put number type values in row 2, 3, 4 for Column B
			cells["B2"].PutValue(220000);
			cells["B3"].PutValue(80000);
			cells["B4"].PutValue(150000);

            //Put number type values in row 2, 3, 4 for Column C
			cells["C2"].PutValue(100000);
			cells["C3"].PutValue(150000);
			cells["C4"].PutValue(60000);			
		}

		private void CreateCellsFormatting(Workbook workbook)
		{
            //Initialize Style1
			Style style1 = workbook.Styles[workbook.Styles.Add()];
			
            //Set border settings for style1
			style1.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
			
            //Set Font property IsBold to False
            style1.Font.IsBold = true;

            //Set alignment for Style1
			style1.HorizontalAlignment = TextAlignmentType.Center;			

            //intialize Cells
			Cells cells = workbook.Worksheets[0].Cells;
			
            //Apply Style1 on A1, B1,C1
            cells["A1"].SetStyle(style1);
			cells["B1"].SetStyle(style1);
			cells["C1"].SetStyle(style1);
			

            //Intialize style2
			Style style2 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy data from another style object
		    style2.Copy(style1);
			
            //Set isBold property of style Font to False
            style2.Font.IsBold = false;

            //Set alignment of style2
			style2.HorizontalAlignment = TextAlignmentType.Right;
			
            //Set foreground color
			style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
			
            //Set Pattern of Style2
            style2.Pattern = BackgroundType.Solid;	

            //Apply style2 to A2 and A4
			cells["A2"].SetStyle(style2);
			cells["A4"].SetStyle(style2);

            //Initialize Style3
			Style style3 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy properties from Style2
            style3.Copy(style2);

			//Set cell format
			style3.Custom = "\"$\"#,##0";

            //Apply Style to B2, C2, B4 and C4
			cells["B2"].SetStyle(style3);
			cells["C2"].SetStyle(style3);
			cells["B4"].SetStyle(style3);
			cells["C4"].SetStyle(style3);			
			
            //Initialize Style4
			Style style4 = workbook.Styles[workbook.Styles.Add()];

            //copy contents from Style2
			style4.Copy(style2);

			//Sets foreground color
			style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
			
            //Set style Pattern to solid
            style4.Pattern = BackgroundType.Solid;	

            //Apply Style to A3
			cells["A3"].SetStyle(style4);

            //Initialize Style5
			Style style5 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy properties from Style4
            style5.Copy(style4);

			//Set cell format
			style5.Custom = "\"$\"#,##0";

            //Set Style to B3 and C3
			cells["B3"].SetStyle(style5);
			cells["C3"].SetStyle(style5);
		}		

		private void CreateStaticReport(Workbook workbook)
		{
            //Initialize worksheet
			Worksheet sheet = workbook.Worksheets[0];
			
            //Set the name of worksheet
			sheet.Name = "Clustered Bar";
			
            //Set GridLines invisible
            sheet.IsGridlinesVisible = false;

            //Create chart, If Check box on Ui is Checked then create Bar3DClustered Chart else Bar Chart
			int indexChart = 0;
			if ( checkBoxShow3D.Checked)
				indexChart = sheet.Charts.Add(ChartType.Bar3DClustered,5,1,26,10);	
			else
				indexChart = sheet.Charts.Add(ChartType.Bar,5,1,26,10);	
			Chart chart = sheet.Charts[indexChart];

			//Set properties of chart based upon state of check box on UI
            if ( checkBoxShow3D.Checked)
                chart.PlotArea.Border.IsVisible = false;
			chart.CategoryAxis.MajorGridLines.IsVisible = false;

			//Set properties of chart title
			chart.Title.Text = "Fruit Sales By Region";
			chart.Title.TextFont.Color = Color.Black;
			chart.Title.TextFont.IsBold = true;
			chart.Title.TextFont.Size = 12;

			//Set properties of nseries
			chart.NSeries.Add("B2:C4", true);
			chart.NSeries.CategoryData = "A2:A4";
			chart.NSeries.IsColorVaried = true;

            //Initialize Cells
			Cells cells = workbook.Worksheets[0].Cells;

            //loop over the Chart's Nseries and Assign Name from Cell Values
			for ( int i = 0 ; i < chart.NSeries.Count ; i ++ )
			{
				chart.NSeries[i].Name = cells[0,i+1].Value.ToString();
			}

			//Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Region";
			chart.CategoryAxis.Title.TextFont.Color = Color.Black;
			chart.CategoryAxis.Title.TextFont.IsBold = true;
			chart.CategoryAxis.Title.TextFont.Size = 10;
			chart.CategoryAxis.Title.RotationAngle = 90;					

			//Set properties of legend to show on Top
			chart.Legend.Position = LegendPositionType.Top;
		}

      
		
	}
}
