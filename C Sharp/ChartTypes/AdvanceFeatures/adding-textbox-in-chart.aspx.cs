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
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;


namespace Aspose.Cells.Demos
{
	/// <summary>
    /// Summary description for AddTextBoxInChart.
	/// </summary>
    public class TextBoxInChart : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.CheckBox CheckBoxShow3D;
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
			this.ID = "Area";
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion

        protected void btnProcess_Click(object sender, EventArgs e)
        {
            //Create a new workbook
			Workbook workbook = new Workbook();	

			//Set default font
			Style style = workbook.DefaultStyle;

            //Set font name
			style.Font.Name = "Tahoma";

            //Set workbook default style
			workbook.DefaultStyle = style;	
					
            //Call method to load dummy data in workbook
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
            workbook.Save(HttpContext.Current.Response, "AddTextBoxInChart." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
           	
            // note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticReport(Workbook workbook)
		{
            //Create a new worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            //Adding a sample value to "A1" cell
            worksheet.Cells["A1"].PutValue(50);

            //Adding a sample value to "A2" cell
            worksheet.Cells["A2"].PutValue(100);

            //Adding a sample value to "A3" cell
            worksheet.Cells["A3"].PutValue(150);

            //Adding a sample value to "B1" cell
            worksheet.Cells["B1"].PutValue(60);

            //Adding a sample value to "B2" cell
            worksheet.Cells["B2"].PutValue(32);

            //Adding a sample value to "B3" cell
            worksheet.Cells["B3"].PutValue(50);

            //Adding a chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);

            //Accessing the instance of the newly added chart
            Chart chart = worksheet.Charts[chartIndex];

            //Adding NSeries (chart data source) to the chart ranging from "A1" cell to "B3"
            chart.NSeries.Add("A1:B3", true);

            //Show the Data Table with the chart
            chart.ShowDataTable = true;

            //Add a new textbox to the chart.
            Aspose.Cells.Drawing.TextBox textbox0 = chart.Shapes.AddTextBoxInChart(50, 1200, 350, 2550);

            //Fill the text.
            textbox0.Text = "Sales By Region";

            //Get the textbox text frame.
            Aspose.Cells.Drawing.MsoTextFrame textframe0 = textbox0.TextFrame;

            //Set the textbox to adjust it according to its contents.
            textframe0.AutoSize = true;

            //Set the font color.
            textbox0.Font.Color = Color.Maroon;

            //Set the font to bold.
            textbox0.Font.IsBold = true;
            
            //Set the font size.
            textbox0.Font.Size = 14;
            
            //Set font attribute to italic.
            textbox0.Font.IsItalic = true;

            //Get the filformat of the textbox.
            MsoFillFormat fillformat = textbox0.FillFormat;
            
            //Set the fillcolor.
            fillformat.ForeColor = Color.Silver;
            
            //Get the lineformat type of the textbox.
            MsoLineFormat lineformat = textbox0.LineFormat;
            
            //Set the line style.
            lineformat.Style = MsoLineStyle.ThinThick;

            //Set the line weight.
            lineformat.Weight = 2;

            //Set the dash style to solid.
            lineformat.DashStyle = MsoLineDashStyle.Solid;
		   
		}       
	}
}
