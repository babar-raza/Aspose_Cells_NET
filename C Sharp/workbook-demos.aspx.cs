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
	/// Summary description for WorkbookDemos.
	/// </summary>
	public class WorkbookDemos : System.Web.UI.Page
	{
	
		private void Page_Load(object sender, System.EventArgs e)
		{
			// Put user code to initialize the page here
			/*if (!Page.IsPostBack)
			{
				if (Request.QueryString["Type"]!=null)
				{
					int type = int.Parse(Request.QueryString["Type"].ToString());
					switch (type)
					{
					    //========about Data=====================//
						case 1:	
							HelloWorld.CreateStaticReport();								
							break;
						case 2:
							FindOrSearchData.CreateStaticReport();
							break;
						case 5:	
							DataFilter.CreateStaticReport();												
							break;
						case 6:
							DataValidation.CreateStaticReport();
							break;
					    case 7:
							SetFormula.CreateStaticReport();
							break;
						case 8:
							CalculateFormula.CreateStaticReport();
							break;
						case 9:							
							AddingHyperlinks.CreateStaticReport();
							break;
						case 10:
							NamedRanges.CreateStaticReport();
							break;		
						//=========about Formatting ===============//
						case 11:
							NumberFormatting.CreateStaticReport();
							break;
						case 12:
							AlignmentSetting.CreateStaticReport();
							break;
						case 13:
							FontSetting.CreateStaticReport();
							break;
						case 14:
							BorderSetting.CreateStaticReport();
							break;
						case 15:
							PatternSetting.CreateStaticReport();
							break;
						case 16:
							FormattingRange.CreateStaticReport();
							break;
						//==============about Row/Column settings =========//
						case 20:
							AdjustingRowsAndColumns.CreateStaticReport();
							break;
					    case 21:
							AutoFitRowsAndColumns.CreateStaticReport();
							break;
						//===============about Worksheet Setting ==========//
						case 30:
							FreezePanes.CreateStaticReport();
							break;
						//==============about Drawing Objects ============//
						case 31:
							AddingPictures.CreateStaticReport();
							break;
						case 32:
							AddingComments.CreateStaticReport();
							break;
						case 33:
							OtherDrawingObjects.CreateStaticReport();
							break;
						//==============about Security settings ============//
						case 34:
							ProtectingWorksheet.CreateStaticReport();
							break;
						case 35:
							AdvancedProtection.CreateStaticReport();
							break;
						case 36:
							UnprotectAWorksheet.CreateStaticReport();
							break;
						//=============about Setting Page Options ==========//
						case 37:
							SettingPageOption.CreateStaticReport();
							break;
						case 38:
							SettingMargins.CreateStaticReport();
							break;
						case 39:
							HeadersAndFooters.CreateStaticReport();
							break;
						case 40:
							SettingPrintOptions.CreateStaticReport();
							break;
					}
				}
			}	*/		
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
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion
	}
}
