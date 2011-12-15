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
    /// Summary description for ImageFillFormat.
	/// </summary>
    public partial class ImageFillFormat : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.CheckBox CheckBoxShow3D;
	
		protected void Page_Load(object sender, System.EventArgs e)
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
			this.ID = "Area";

		}
		#endregion

		protected void btnProcess_Click(object sender, EventArgs e)
        {
			Workbook workbook = new Workbook();	

			//Set default font
			Style style = workbook.DefaultStyle;
			style.Font.Name = "Tahoma";
			workbook.DefaultStyle = style;	
					
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
            workbook.Save(HttpContext.Current.Response, "ImageFillFormat." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			
            // note by Vit - end response to avoid unneeded html after xls
            Response.End();
		}

		private void CreateStaticReport(Workbook workbook)
		{
            //Get First Worksheet of the Workbook
            Worksheet ws = workbook.Worksheets[0];

            //Adding data to cells
            Cells cells = ws.Cells;
            
            //Insert cell String contents in Column A
            cells["A1"].PutValue("Aspose.Cells");

            cells["A2"].PutValue("Aspose.Words");

            cells["A3"].PutValue("Aspose.PDF");

            //Insert Cell number contents in Column B
            cells["B1"].PutValue(35);

            cells["B2"].PutValue(50);

            cells["B3"].PutValue(15);


            //Create a Pie Type Chart in worksheet charts collection
            int index = ws.Charts.Add(ChartType.Pie, 4, 1, 30, 10);

            Chart chart = ws.Charts[index];

            //Assign range of cells as charts N-Series
            chart.NSeries.Add("B1:B3", true);

            //define N-series Category Data
            chart.NSeries.CategoryData = "A1:A3";

            //Set Datalabels
            chart.NSeries[0].DataLabels.ShowPercentage = true;

            
            //Create a Stream object AND intialize it with path to Image
            FileStream fstream = new FileStream(System.Web.HttpContext.Current.Server.MapPath("~/Image/school.jpg"), FileMode.Open);            

            //Read Byte Data into any Array
            byte[] ImageData = new Byte[fstream.Length];

            //Obtain the file into the array of bytes from streams.
            fstream.Read(ImageData, 0, ImageData.Length);

            //Fillformat as Image
            chart.ChartArea.Area.FillFormat.ImageData = ImageData;
            
            ws.AutoFitColumns();
            fstream.Close();   
		}		
	}
}
