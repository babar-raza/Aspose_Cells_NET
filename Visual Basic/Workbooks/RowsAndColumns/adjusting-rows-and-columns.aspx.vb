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

Partial Public Class Workbooks_RowsAndColumns_AdjustingRowsAndColumns
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()

		Dim workbook As New Workbook()

		Dim cells As Cells = workbook.Worksheets(0).Cells

		'Set the height of all row in the worksheet
		cells.StandardHeight = 20

		'Set the width of all columns in the worksheet
		cells.StandardWidth = 20

		'Set the width of the first column 
		cells.SetColumnWidth(0, 12)

		'Set the width of the column 
		cells.SetColumnWidth(1, 40)
		'Set the height of the row 
		cells.SetRowHeight(1, 8)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "RowHeightandColumnWidth.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "RowHeightandColumnWidth.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub

End Class
