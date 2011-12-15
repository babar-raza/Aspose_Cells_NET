Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Web
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports Aspose.Cells

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for WorkbookDemos.
	''' </summary>
	Public Class WorkbookDemos
		Inherits System.Web.UI.Page

		Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			' Put user code to initialize the page here
'            if (!Page.IsPostBack)
'			{
'				if (Request.QueryString["Type"]!=null)
'				{
'					int type = int.Parse(Request.QueryString["Type"].ToString());
'					switch (type)
'					{
'                        //========about Data=====================//
'						case 1:	
'							HelloWorld.CreateStaticReport();								
'							break;
'						case 2:
'							FindOrSearchData.CreateStaticReport();
'							break;
'						case 5:	
'							DataFilter.CreateStaticReport();												
'							break;
'						case 6:
'							DataValidation.CreateStaticReport();
'							break;
'					    case 7:
'							SetFormula.CreateStaticReport();
'							break;
'						case 8:
'							CalculateFormula.CreateStaticReport();
'							break;
'						case 9:							
'							AddingHyperlinks.CreateStaticReport();
'							break;
'						case 10:
'							NamedRanges.CreateStaticReport();
'							break;		
'                        //=========about Formatting ===============//
'						case 11:
'							NumberFormatting.CreateStaticReport();
'							break;
'						case 12:
'							AlignmentSetting.CreateStaticReport();
'							break;
'						case 13:
'							FontSetting.CreateStaticReport();
'							break;
'						case 14:
'							BorderSetting.CreateStaticReport();
'							break;
'						case 15:
'							PatternSetting.CreateStaticReport();
'							break;
'						case 16:
'							FormattingRange.CreateStaticReport();
'							break;
'                        //==============about Row/Column settings =========//
'						case 20:
'							AdjustingRowsAndColumns.CreateStaticReport();
'							break;
'					    case 21:
'							AutoFitRowsAndColumns.CreateStaticReport();
'							break;
'                        //===============about Worksheet Setting ==========//
'						case 30:
'							FreezePanes.CreateStaticReport();
'							break;
'                        //==============about Drawing Objects ============//
'						case 31:
'							AddingPictures.CreateStaticReport();
'							break;
'						case 32:
'							AddingComments.CreateStaticReport();
'							break;
'						case 33:
'							OtherDrawingObjects.CreateStaticReport();
'							break;
'                        //==============about Security settings ============//
'						case 34:
'							ProtectingWorksheet.CreateStaticReport();
'							break;
'						case 35:
'							AdvancedProtection.CreateStaticReport();
'							break;
'						case 36:
'							UnprotectAWorksheet.CreateStaticReport();
'							break;
'                        //=============about Setting Page Options ==========//
'						case 37:
'							SettingPageOption.CreateStaticReport();
'							break;
'						case 38:
'							SettingMargins.CreateStaticReport();
'							break;
'						case 39:
'							HeadersAndFooters.CreateStaticReport();
'							break;
'						case 40:
'							SettingPrintOptions.CreateStaticReport();
'							break;
'					}
'				}
'			}			
		End Sub

		#Region "Web Form Designer generated code"
		Overrides Protected Sub OnInit(ByVal e As EventArgs)
			'
			' CODEGEN: This call is required by the ASP.NET Web Form Designer.
			'
			InitializeComponent()
			MyBase.OnInit(e)
		End Sub

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region
	End Class
End Namespace
