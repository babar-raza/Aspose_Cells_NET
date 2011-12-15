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
	''' Summary description for GroupingRowsAndColumns.
	''' </summary>
	Public Class GroupingRowsAndColumns
		Inherits System.Web.UI.Page
		Protected WithEvents Button1 As System.Web.UI.WebControls.Button
		Protected WithEvents Button2 As System.Web.UI.WebControls.Button

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
'			Me.Button1.Click += New System.EventHandler(Me.Button1_Click);
'			Me.Button2.Click += New System.EventHandler(Me.Button2_Click);
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region

		Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
			'Open template
			Dim path As String = MapPath("~")
			path = path.Substring(0, path.LastIndexOf("\"))
			path &= "\designer\Workbooks\GroupingRowsAndColumns.xls"


			Dim workbook As New Workbook(path)

			GroupRowsAndColumns(workbook)

			workbook.Save(HttpContext.Current.Response, "GroupingRowsAndColumns.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))

			' End response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
			'Open template
			Dim path As String = MapPath("~")
			path = path.Substring(0, path.LastIndexOf("\"))
			path &= "\designer\Workbooks\UnGroupingRowsAndColumns.xls"

			Dim workbook As New Workbook(path)

			UnGroupRowsAndColumns(workbook)

			workbook.Save(HttpContext.Current.Response, "UnGroupingRowsAndColumns.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		End Sub

		Private Sub GroupRowsAndColumns(ByVal workbook As Workbook)
			Dim worksheet As Worksheet = workbook.Worksheets(0)
			worksheet.Cells.GroupRows(0, 9)
			worksheet.Cells.GroupColumns(0, 1)

			'Set SummaryRowBelow property
			worksheet.Outline.SummaryRowBelow = True

			'Set SummaryColumnRight property
			worksheet.Outline.SummaryColumnRight = False
		End Sub

		Private Sub UnGroupRowsAndColumns(ByVal workbook As Workbook)
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			worksheet.Cells.UngroupRows(0, 9)
			worksheet.Cells.UngroupColumns(0, 1)
		End Sub
	End Class
End Namespace
