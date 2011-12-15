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
Imports System.Drawing

Partial Public Class Workbooks_Formatting_BorderSetting
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
		'Get the cells collection in the first worksheet
		Dim cells As Cells = workbook.Worksheets(0).Cells

		'Get Style of B2
		Dim style As Aspose.Cells.Style = cells("B2").GetStyle()

		'Set the cell border color
		style.Borders(BorderType.TopBorder).Color = Color.Blue
		style.Borders(BorderType.BottomBorder).Color = Color.Blue
		style.Borders(BorderType.LeftBorder).Color = Color.Blue
		style.Borders(BorderType.RightBorder).Color = Color.Blue
		style.Borders(BorderType.DiagonalDown).Color = Color.Blue
		style.Borders(BorderType.DiagonalUp).Color = Color.Blue


		'Set the cell border type
		style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.DashDot
		style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.DashDot
		style.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.DashDot
		style.Borders(BorderType.RightBorder).LineStyle = CellBorderType.DashDot
		style.Borders(BorderType.DiagonalDown).LineStyle = CellBorderType.DashDot
		style.Borders(BorderType.DiagonalUp).LineStyle = CellBorderType.DashDot

		'Setting Border Style for B2
		cells("B2").SetStyle(style)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "BorderSetting.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "BorderSetting.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class



