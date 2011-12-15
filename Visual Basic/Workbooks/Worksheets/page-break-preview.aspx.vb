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
	''' Summary description for PageBreakPreview.
	''' </summary>
	Public Class PageBreakPreview
		Inherits System.Web.UI.Page
		Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList
		Protected WithEvents Button2 As System.Web.UI.WebControls.Button
		Protected WithEvents Button1 As System.Web.UI.WebControls.Button

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
			'Create a new workbook
			Dim workbook As New Workbook()

			'Get the first worksheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Display the worksheet in page break preview
			worksheet.IsPageBreakPreview = True

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "DisplayPageBreakPreview.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "DisplayPageBreakPreview.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub

		Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
			'Create a new workbook
			Dim workbook As New Workbook()

			'Get the first worksheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Hide the worksheet in page break preview
			worksheet.IsPageBreakPreview = False

			'Save the excel file
			workbook.Save(HttpContext.Current.Response, "HidePageBreakPreview.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))

			' End response to avoid unneeded html after xls
			Response.End()

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "HidePageBreakPreview.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "HidePageBreakPreview.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub
	End Class
End Namespace


