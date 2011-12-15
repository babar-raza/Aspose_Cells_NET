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
	''' Summary description for PageBreaks.
	''' </summary>
	Public Class PageBreaks
		Inherits System.Web.UI.Page
		Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList
		Protected WithEvents Button2 As System.Web.UI.WebControls.Button
		Protected WithEvents Button3 As System.Web.UI.WebControls.Button
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
'			Me.Button3.Click += New System.EventHandler(Me.Button3_Click);
'			Me.Button2.Click += New System.EventHandler(Me.Button2_Click);
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region

		Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
			Dim workbook As New Workbook()
			Dim sheet As Worksheet = workbook.Worksheets(0)

			CreateStaticData(workbook)

			'Add a page break at cell B2
			sheet.HorizontalPageBreaks.Add("B2")
			sheet.VerticalPageBreaks.Add("B2")

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "AddPageBreaks.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "AddPageBreaks.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			Dim cells As Cells = workbook.Worksheets(0).Cells
			'Put a value into a cell
			cells("A1").PutValue("World")
			cells("A2").PutValue("Aspose")
			cells("A3").PutValue(100)
			cells("B1").PutValue(200)
			cells("B2").PutValue(300)
			cells("B3").PutValue(500)
		End Sub

		Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
			Dim workbook As New Workbook()
			Dim sheet As Worksheet = workbook.Worksheets(0)

			CreateStaticData(workbook)

			'Add a page break at cell B2
			sheet.HorizontalPageBreaks.Add("B2")
			sheet.VerticalPageBreaks.Add("B2")
			sheet.HorizontalPageBreaks.Add(5, 1)
			sheet.HorizontalPageBreaks.Add(6, 1, 10)
			'Remove a page break at cell 
			sheet.HorizontalPageBreaks.RemoveAt(0)
			sheet.VerticalPageBreaks.RemoveAt(0)

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "RemovetPageBreaks.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "RemovetPageBreaks.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()

		End Sub

		Private Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
			Dim workbook As New Workbook()

			CreateStaticData(workbook)

			'Clear all page breaks
			workbook.Worksheets(0).HorizontalPageBreaks.Clear()
			workbook.Worksheets(0).VerticalPageBreaks.Clear()

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "ClearPageBreaks.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "ClearPageBreaks.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub
	End Class
End Namespace
