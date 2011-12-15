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
Imports System.IO

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for ExportData.
	''' </summary>
	Public Class ExportData
		Inherits System.Web.UI.Page
		Protected dgExportData As System.Web.UI.WebControls.DataGrid
		Protected WithEvents btnExportData As System.Web.UI.WebControls.Button

		Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			' Put user code to initialize the page here
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
'			Me.btnExportData.Click += New System.EventHandler(Me.btnExportData_Click);
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region

		Private Sub btnExportData_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExportData.Click
			'Open template
			Dim path As String = MapPath("~")
			path = path.Substring(0, path.LastIndexOf("\"))
			path &= "\designer\book1.xls"

			'Instantiate a new workbook
			Dim workbook As New Workbook(path)

			'Get the first worksheet in the workbook
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Create a datatable
			Dim dataTable As New DataTable()

			'Export worksheet data to a DataTable object by calling either ExportDataTable or ExportDataTableAsString method of the Cells class		 	
			dataTable = worksheet.Cells.ExportDataTable(0, 0, worksheet.Cells.MaxRow + 1, worksheet.Cells.MaxColumn + 1)

			'Bind the DataGrid with DataTable
			dgExportData.DataSource = dataTable
			dgExportData.ShowHeader = False
			dgExportData.DataBind()

		End Sub

	End Class
End Namespace


