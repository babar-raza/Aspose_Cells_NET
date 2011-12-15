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
    /// Summary description
    /// </summary>
    public class SettingBackgroundImageOfChartSheet : System.Web.UI.Page
    {
        protected System.Web.UI.WebControls.Button Button2;
        protected System.Web.UI.WebControls.Button Button1;
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
            InitializeComponent();
            base.OnInit(e);
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Button1.Click += new System.EventHandler(this.Button1_Click);
            
            this.Load += new System.EventHandler(this.Page_Load);

        }
        #endregion

        private void Button1_Click(object sender, System.EventArgs e)
        {
            //Create a new workbook
            Workbook workbook = new Workbook();

            AddWorksheets(workbook);
            
            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "ChartSheet.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "ChartSheet.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End();  
        }

        private void AddWorksheets(Workbook workbook)
        {
            ////Create a Stream object
            FileStream fstream = new FileStream(System.Web.HttpContext.Current.Server.MapPath("~/Image/school.JPG"), FileMode.Open);

            byte[] Data = new Byte[fstream.Length];

            ////Obtain the file into the array of bytes from streams.
            fstream.Read(Data, 0, Data.Length);

            //Get First Worksheet of the Workbook
            Worksheet ws = workbook.Worksheets[0];

            //Set Worksheet Type
            ws.Type = SheetType.Chart;

            //Set Worksheet background image
            ws.SetBackground(Data);

            //Add new Data Sheet
            Worksheet data = workbook.Worksheets.Add("Sheet2");

            //Get data sheet's cells collection
            Cells cells = data.Cells;

            //Add Values to cells
            cells["A1"].PutValue("Aspose.Cells");

            cells["A2"].PutValue("Aspose.Words");

            cells["A3"].PutValue("Aspose.PDF");

            cells["B1"].PutValue(35);

            cells["B2"].PutValue(55);

            cells["B3"].PutValue(10);

            //Adding a new chart
            int index = ws.Charts.Add(ChartType.Pie, 5, 0, 15, 5);

            //get newly added chart
            Chart chart = ws.Charts[index];

            //add nseries of the chart
            chart.NSeries.Add("Sheet2!B1:B3", true);
            chart.NSeries.CategoryData = "Sheet2!A1:A3";

            //show Data Labels
            chart.NSeries[0].DataLabels.ShowCategoryName = true;
            chart.NSeries[0].DataLabels.ShowPercentage = true;
            
            //No formatting for Chart Area
            chart.ChartArea.Area.Formatting = FormattingType.None;
        }             
       
    }
}


