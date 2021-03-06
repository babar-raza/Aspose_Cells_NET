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
    /// Summary description for DisplayHideScrollBars.
    /// </summary>
    public class DisplayHideScrollBars : System.Web.UI.Page
    {
        protected System.Web.UI.WebControls.Button Button1;
        protected System.Web.UI.WebControls.Button Button2;
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
            this.Button2.Click += new System.EventHandler(this.Button2_Click);
            this.Load += new System.EventHandler(this.Page_Load);

        }
        #endregion

        private void Button1_Click(object sender, System.EventArgs e)
        {
            //Create a new workbook
            Workbook workbook = new Workbook();

            //Set a value indicating whether the generated spreadsheet will contain a horizontal scroll bar
            workbook.Settings.IsHScrollBarVisible = true;
            
            //Set a value indicating whether the generated spreadsheet will contain a vertical scroll bar
            workbook.Settings.IsVScrollBarVisible = true;
            
            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "DisplayScrollBars.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "DisplayScrollBars.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End();   
        }

        private void Button2_Click(object sender, System.EventArgs e)
        {
            //Create a new workbook
            Workbook workbook = new Workbook();

            //Set a value indicating whether the generated spreadsheet will contain a horizontal scroll bar
            workbook.Settings.IsHScrollBarVisible = false;
            
            //Set a value indicating whether the generated spreadsheet will contain a vertical scroll bar
            workbook.Settings.IsVScrollBarVisible = false;
            
            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "HideScrollBars.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "HideScrollBars.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End();   
        }
    }
}


