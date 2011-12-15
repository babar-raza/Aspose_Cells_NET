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
	''' Summary description for ManageWorksheets.
	''' </summary>
	Public Class ManagingWorksheets
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

			AddWorksheets(workbook)

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "AddWorksheets.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "AddWorksheets.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub

		Private Sub AddWorksheets(ByVal workbook As Workbook)
			'Get the first worksheet in the workbook
			Dim worksheet As Worksheet = workbook.Worksheets(0)
			'Set the name of the sheet
			worksheet.Name = "My Worksheet1"

			'Add a new worksheet to the Workbook object
			workbook.Worksheets.Add()
			'Obtain the reference of the newly added worksheet by passing its sheet index
			worksheet = workbook.Worksheets(1)
			'Set the name of the newly added worksheet
			worksheet.Name = "My Worksheet2"

			'Add a new worksheet to the Workbook object
			workbook.Worksheets.Add()
			'Obtain the reference of the newly added worksheet by passing its sheet index
			worksheet = workbook.Worksheets(2)
			'Set the name of the newly added worksheet
			worksheet.Name = "My Worksheet3"
		End Sub

		Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click

			'Open template
			Dim path As String = MapPath("~")
			path = path.Substring(0, path.LastIndexOf("\"))
			path &= "\designer\Workbooks\ManagingWorksheets.xls"
			Dim workbook As New Workbook(path)

			'Remove a worksheet from an Excel file using its sheet index
			workbook.Worksheets.RemoveAt(1)

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "RemoveWorksheets.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "RemoveWorksheets.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub
	End Class
End Namespace


