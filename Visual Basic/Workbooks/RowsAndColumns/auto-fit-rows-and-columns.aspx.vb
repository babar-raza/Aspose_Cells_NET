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

Partial Public Class Workbooks_RowsAndColumns_AutoFitRowsAndColumns
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		Dim workbook As New Workbook()
		Dim sheet As Worksheet = workbook.Worksheets(0)

		Dim cells As Cells = sheet.Cells

		cells("B1").PutValue("Aspose.Cells")
		'Get Style Object 
		Dim style As Aspose.Cells.Style = cells("B1").GetStyle()

		style.RotationAngle = 45
		style.Font.IsBold = True
		cells("B1").SetStyle(style)

		'Auto row fit
		sheet.AutoFitRow(0)
		'Auto column fit
		sheet.AutoFitColumn(1)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "AutoFitRowsAndColumns.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "AutoFitRowsAndColumns.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
