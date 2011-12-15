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

Partial Public Class Workbooks_Formatting_NumberFormatting
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()

		'Open template
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
		path = path.Substring(0, path.LastIndexOf("\"))
		path &= "\designer\Workbooks\NumberFormatting.xls"


		'Create a new workbook
		Dim workbook As New Workbook(path)

		'Get the cells collection in the workbook
		Dim cells As Cells = workbook.Worksheets(0).Cells

		Dim style As Aspose.Cells.Style

		'Set number format with built-in index
		For i As Integer = 1 To 36
			cells(i, 1).PutValue(1234.5)

			Dim Number As Integer = cells(i, 0).IntValue

			'Get Style of Cell
			style = cells(i, 1).GetStyle()

			'Set the display number format
			style.Number = Number

			'Apply Style
			cells(i, 1).SetStyle(style)
		Next i

		'Set number format with custom format string
		For i As Integer = 1 To 3
			cells(i, 3).PutValue(1234.5)

			'Get Style of Cell
			style = cells(i, 3).GetStyle()

			'Set the display custom number format
			style.Custom = cells(i, 2).StringValue

			'Apply Style
			cells(i, 3).SetStyle(style)
		Next i

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "NumberFormatting.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "NumberFormatting.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub

End Class



