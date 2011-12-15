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

Partial Public Class Workbooks_Data_NamedRanges
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		'Call Method to create report
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Create a new workbook
		Dim workbook As New Workbook()

		'Get the first worksheet in the workbook
		Dim sheet As Worksheet = workbook.Worksheets(0)

		'Get the cells collection in the sheet
		Dim cells As Cells = sheet.Cells

		'Create a named range
		Dim range As Range = cells.CreateRange("B1", "E5")

		'Set the name of the named range
		range.Name = "TestRange"

		'Accessing a specific Named Range
		Dim myRange As Range = workbook.Worksheets.GetRangeByName("TestRange")

		'Get the first cell in the range
		Dim cell As Aspose.Cells.Cell = myRange(0, 0)

		'Put string value to it
		cell.PutValue("Top left of TestRange")

		'Get Style Object 
		Dim style As Aspose.Cells.Style = cell.GetStyle()

		'Set the fill color of the cell
		style.ForegroundColor = System.Drawing.Color.Blue
		style.Pattern = BackgroundType.Solid
		cell.SetStyle(style)

		'Get the last cell in the range
		cell = myRange(myRange.RowCount - 1, myRange.ColumnCount - 1)

		'Put a string value to it
		cell.PutValue("Bottom right of TestRange")

		'Set the fill color of the cell
		style.ForegroundColor = System.Drawing.Color.Blue
		style.Pattern = BackgroundType.Solid
		cell.SetStyle(style)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "NamedRanges.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "NamedRanges.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub
End Class

