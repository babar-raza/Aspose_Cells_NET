using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace Aspose.Cells.Demos
{
    public partial class UsingSparklines : System.Web.UI.Page
    {
        protected System.Web.UI.WebControls.DropDownList ddlFileVersion;

        protected void Page_Load(object sender, EventArgs e)
        {

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
            //this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            this.Load += new System.EventHandler(this.Page_Load);

        }
        #endregion
        protected void btnProcess_Click(object sender, EventArgs e)
        {
            CreateStaticReport();
        }

        protected void CreateStaticReport()
        {
            //Intiaalize workbook with xlsx file format
            Workbook workbook = new Workbook();

            //Clear workbook's worksheets
            workbook.Worksheets.Clear();

            //Insert new Worksheet in workbook and name it "New"
            Worksheet worksheet = workbook.Worksheets.Add("New");

            //Insert dummy data in A8, A9 and A10 cells
            worksheet.Cells["A8"].PutValue(34);
            worksheet.Cells["A9"].PutValue(50);
            worksheet.Cells["A10"].PutValue(34);

            //Intialize Cell Area
            CellArea cellArea = new CellArea();

            //Assign Cell Area boundaries
            cellArea.StartColumn = 0;
            cellArea.EndColumn = 0;
            cellArea.StartRow = 0;
            cellArea.EndRow = 0;

            //Add new Sparklines in worksheet's sparlines collection and Assign the area for it
            int index = worksheet.SparklineGroupCollection.Add(SparklineType.Column, worksheet.Name + "!A8:A10", true, cellArea);

            //Initalize Sparklines Group
            SparklineGroup group = worksheet.SparklineGroupCollection[index];


            // change the color of the series if need
            CellsColor cellColor = workbook.CreateCellsColor();
            cellColor.Color = Color.Orange;

            //Asign the group series color
            group.SeriesColor = cellColor;

            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "SparkLines.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "SparkLines.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }
            
            //end response to avoid unneeded html
            HttpContext.Current.Response.End();
        }
    }
}
