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
	''' Summary description for InsertingAndDeletingRowsAndColumns.
	''' </summary>
	Public Class InsertingAndDeletingRowsAndColumns
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
			Dim workbook As New Workbook()
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put values into a cell
			cells("A1").PutValue("1st Row & Column")
			cells("A2").PutValue("2nd Row")
			cells("A3").PutValue("3rd Row")
			cells("A4").PutValue("4th Row")
			cells("A5").PutValue("5th Row")
			cells("A6").PutValue("6th Row")
			cells("A7").PutValue("7th Row")
			cells("A8").PutValue("8th Row")
			cells("A9").PutValue("9th Row")
			cells("A10").PutValue("10th Row")
			cells("A11").PutValue("11th Row")
			cells("A12").PutValue("12th Row")
			cells("A13").PutValue("13th Row")
			cells("A14").PutValue("14th Row")

			cells("B1").PutValue("2nd Column")
			cells("C1").PutValue("3rd Column")
			cells("D1").PutValue("4th Column")
			cells("E1").PutValue("5th Column")

			sheet.AutoFitColumns()

			'Insert 10 rows from the 3rd row
			sheet.Cells.InsertRows(2, 10)

			'Insert 3rd column 
			sheet.Cells.InsertColumn(2)

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "InsertRowsAndColumns.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "InsertRowsAndColumns.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub

		Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
			Dim workbook As New Workbook()
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim cells As Cells = workbook.Worksheets(0).Cells
			'Put a value into a cell
			cells("A1").PutValue("1st Row & Column")
			cells("A2").PutValue("2nd Row")
			cells("A3").PutValue("3rd Row")
			cells("A4").PutValue("4th Row")
			cells("A5").PutValue("5th Row")
			cells("A6").PutValue("6th Row")
			cells("A7").PutValue("7th Row")
			cells("A8").PutValue("8th Row")
			cells("A9").PutValue("9th Row")
			cells("A10").PutValue("10th Row")
			cells("A11").PutValue("11th Row")
			cells("A12").PutValue("12th Row")
			cells("A13").PutValue("13th Row")
			cells("A14").PutValue("14th Row")

			cells("B1").PutValue("2nd Column")
			cells("C1").PutValue("3rd Column")
			cells("D1").PutValue("4th Column")
			cells("E1").PutValue("5th Column")

			sheet.AutoFitColumns()

			'Delete 10 rows from the 3rd row
			sheet.Cells.DeleteRows(2,10)

			'Delete 3rd column
			sheet.Cells.DeleteColumn(2)

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "DeleteRowsAndColumns.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "DeleteRowsAndColumns.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub
	End Class
End Namespace
