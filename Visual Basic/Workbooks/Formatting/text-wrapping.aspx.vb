Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Configuration
Imports System.Collections
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports Aspose.Cells

Partial Public Class Workbooks_Formatting_TextWrapping
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Create Workbook Object
		Dim wb As New Workbook()

		'Open first Worksheet in the workbook
		Dim ws As Worksheet = wb.Worksheets(0)

		'Get Worksheet Cells Collection
		Dim cell As Aspose.Cells.Cells = ws.Cells

		'Increase the width of First Column Width
		cell.SetColumnWidth(0, 35)

		'Increase the height of first row
		cell.SetRowHeight(0, 36)

		'Add Text to the Firts Cell
		cell(0, 0).PutValue("This is the example of text wrap functionality using Aspose.Cells.")

		'Get Style
		Dim style As Aspose.Cells.Style = cell(0, 0).GetStyle()

		'Make Cell's Text wrap
		style.IsTextWrapped = True

		'Set Style
		cell(0, 0).SetStyle(style)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			wb.Save(HttpContext.Current.Response, "TextWrapping.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			wb.Save(HttpContext.Current.Response, "TextWrapping.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub

End Class



