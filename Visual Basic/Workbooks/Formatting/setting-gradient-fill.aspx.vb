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
Imports System.Drawing
Imports Aspose.Cells

Partial Public Class Workbooks_Formatting_GradientFill
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Create a new workbook
		Dim workbook As New Workbook()

		'Get first worksheet in the workbook
		Dim sheet As Worksheet = workbook.Worksheets(0)

		'Get cell A1 from worksheet's cell collection
		Dim cell As Aspose.Cells.Cell = sheet.Cells("A1")

		'Get style of the cell
		Dim style As Aspose.Cells.Style = cell.GetStyle()

		'Set Two Color Gradient
		style.SetTwoColorGradient(Color.Red, Color.Green, Aspose.Cells.Drawing.GradientStyleType.Horizontal, 1)

		'Apply cell style
		cell.SetStyle(style)

		'Set row height and column width
		sheet.Cells.SetColumnWidth(0, 50)
		sheet.Cells.SetRowHeight(0, 50)


		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "GradientFill.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "GradientFill.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class



