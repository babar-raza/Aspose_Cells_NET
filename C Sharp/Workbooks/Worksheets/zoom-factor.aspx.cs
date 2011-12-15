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

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for ZoomFactor.
    /// </summary>
    public class ZoomFactor : System.Web.UI.Page
    {
        protected System.Web.UI.WebControls.DropDownList ddlFileVersion;
        protected System.Web.UI.WebControls.DropDownList Zoom;
        protected System.Web.UI.WebControls.Button Button1;

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
            
            //Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            //Set the zoom factor of the worksheet 
            worksheet.Zoom = int.Parse(Zoom.SelectedItem.Text);

            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "ZoomFactor.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "ZoomFactor.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End();  
        }
    }
}
