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
	/// Summary description for ClusteredColumn.
	/// </summary>
	public class ClusteredColumn : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.TextBox CategoryAxisTitle;
		protected System.Web.UI.WebControls.TextBox ValueAxisTitle;
		protected System.Web.UI.WebControls.DropDownList ValueMaxValue;
		protected System.Web.UI.WebControls.DropDownList ValueMinValue;
		protected System.Web.UI.WebControls.DropDownList ValueMajorUnit;
		protected System.Web.UI.WebControls.DropDownList ValueMinorUnit;
		protected System.Web.UI.WebControls.DropDownList GapWidth;
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
            workbook.Save(HttpContext.Current.Response, "ClusteredColumn." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			// note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}	

		private void CreateStaticData(Workbook workbook)
		{
			Cells cells = workbook.Worksheets[0].Cells;
			//Put string values for cells in first column
			cells["A1"].PutValue("Region");
			cells["A2"].PutValue("France");
			cells["A3"].PutValue("Germany");
			cells["A4"].PutValue("England");
			
            //Put Number values for cells in second column
			cells["B1"].PutValue("Marketing Costs");
			cells["B2"].PutValue(70000);
			cells["B3"].PutValue(55000);
			cells["B4"].PutValue(30000);
		}

		private void CreateCellsFormatting(Workbook workbook)
		{
            //Initialize Style1
			Style style1 = workbook.Styles[workbook.Styles.Add()];
			
            //Set border for Style1
			style1.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
			style1.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
			style1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
			
            //Set Font property IsBold to Fase
            style1.Font.IsBold = true;

            //Set Style Alignments
			style1.HorizontalAlignment = TextAlignmentType.Center;
			style1.VerticalAlignment = TextAlignmentType.Center;

            //Initalize Cells
			Cells cells = workbook.Worksheets[0].Cells;
			
            //Set the width of the specified column
			cells.SetColumnWidth(1,15);
			cells.SetColumnWidth(1,15);

			//Apply style on cells A1 and B1
            cells["A1"].SetStyle(style1);
			cells["B1"].SetStyle(style1);
			
			//Initialize Style2
			Style style2 = workbook.Styles[workbook.Styles.Add()];
			
            //Copy data from another style object
			style2.Copy(style1);
			
            //Set Font Property IsBold to False
            style2.Font.IsBold = false;

            //Set Style Alignmet
			style2.HorizontalAlignment = TextAlignmentType.Right;

			//Set foreground color
			style2.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
			
            //Set Style Patern
            style2.Pattern = BackgroundType.Solid;
			
            //Apply Style of A2 and A4
            cells["A2"].SetStyle(style2);
			cells["A4"].SetStyle(style2);

            //Initialize Style3
			Style style3 = workbook.Styles[workbook.Styles.Add()];			
			
            //Copy the properties from Style2
            style3.Copy(style2);

			//Set cell format
			style3.Custom = "\"$\"#,##0";
	
			//Apply Style to cells B2 and B4 
            cells["B2"].SetStyle(style3);
			cells["B4"].SetStyle(style3);			

            //Initialize Style4
			Style style4 = workbook.Styles[workbook.Styles.Add()];			

            //Copy the properties of Style2
			style4.Copy(style2);

			//Sets foreground color
			style4.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
			
            //Set Style Pattern
            style4.Pattern = BackgroundType.Solid;	
			
            //Apply Style to cell A3
            cells["A3"].SetStyle(style4);

            //Initialize Style5
			Style style5 = workbook.Styles[workbook.Styles.Add()];			
			
            //Copy properties from Style4
            style5.Copy(style4);

			//Set cell format
			style5.Custom = "\"$\"#,##0";

			//Apply Stype to B3
            cells["B3"].SetStyle(style5);
		}

        private void CreateStaticReport(Workbook workbook)
        {
            //Initialize Worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set the name of the worksheet
            sheet.Name = "Clustered Column";

            //Set Gridlines to Invisible
            sheet.IsGridlinesVisible = false;

            //Create chart at next index on worksheet's chart collection
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 1, 29, 10);
           
            //Initialize Chart
            Chart chart = sheet.Charts[chartIndex];

            //Add the nseries collection to a chart 
            chart.NSeries.Add("B2:B4", true);
            
            //Get or set the range of category axis values
            chart.NSeries.CategoryData = "A2:A4";
            chart.NSeries.IsColorVaried = true;
            
            //Loop over the NSeries and Set DataLabels to Show Value
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                chart.NSeries[i].DataLabels.ShowValue = true;
            }

            //Set the legend position to Top
            chart.Legend.Position = LegendPositionType.Top;
            chart.GapWidth = int.Parse(GapWidth.SelectedItem.Text);
            chart.CategoryAxis.MajorGridLines.IsVisible = false;

            //Set properties of chart title
            chart.Title.Text = "Marketing Costs by Region";
            chart.Title.TextFont.IsBold = true;
            chart.Title.TextFont.Color = Color.Black;
            chart.Title.TextFont.Size = 12;

            //Set properties of categoryaxis title
            chart.CategoryAxis.Title.Text = CategoryAxisTitle.Text;
            chart.CategoryAxis.Title.TextFont.Color = Color.Black;
            chart.CategoryAxis.Title.TextFont.IsBold = true;
            chart.CategoryAxis.Title.TextFont.Size = 10;

            //Set properties of valueaxis title
            chart.ValueAxis.Title.Text = ValueAxisTitle.Text;
            chart.ValueAxis.Title.TextFont.Name = "Arial";
            chart.ValueAxis.Title.TextFont.Color = Color.Black;
            chart.ValueAxis.Title.TextFont.IsBold = true;
            chart.ValueAxis.Title.TextFont.Size = 10;
            chart.ValueAxis.Title.RotationAngle = 90;
            chart.ValueAxis.MajorUnit = double.Parse(ValueMajorUnit.SelectedItem.Text);
            chart.ValueAxis.MaxValue = double.Parse(ValueMaxValue.SelectedItem.Text);
            chart.ValueAxis.MinorUnit = double.Parse(ValueMinorUnit.SelectedItem.Text);
            chart.ValueAxis.MinValue = double.Parse(ValueMinValue.SelectedItem.Text);
        }
		
	}
}
