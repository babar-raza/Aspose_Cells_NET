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
using System.IO;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for ExportData.
    /// </summary>
    public class ExportData : System.Web.UI.Page
    {
        protected System.Web.UI.WebControls.DataGrid dgExportData;
        protected System.Web.UI.WebControls.Button btnExportData;

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
            this.btnExportData.Click += new System.EventHandler(this.btnExportData_Click);
            this.Load += new System.EventHandler(this.Page_Load);

        }
        #endregion

        private void btnExportData_Click(object sender, System.EventArgs e)
        {
            //Open template
            string path = MapPath("~");
            path = path.Substring(0, path.LastIndexOf("\\"));
            path += @"\designer\book1.xls";

            //Instantiate a new workbook
            Workbook workbook = new Workbook(path);

            //Get the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            //Create a datatable
            DataTable dataTable = new DataTable();

            //Export worksheet data to a DataTable object by calling either ExportDataTable or ExportDataTableAsString method of the Cells class		 	
            dataTable = worksheet.Cells.ExportDataTable(0, 0, worksheet.Cells.MaxRow + 1,
                         worksheet.Cells.MaxColumn + 1);

            //Bind the DataGrid with DataTable
            dgExportData.DataSource = dataTable;
            dgExportData.ShowHeader = false;
            dgExportData.DataBind();

        }

    }
}


