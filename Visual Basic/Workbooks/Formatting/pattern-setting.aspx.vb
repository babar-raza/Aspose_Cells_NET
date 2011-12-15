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

Partial Public Class Workbooks_Formatting_PatternSetting
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

		'Get the cells collection
		Dim cells As Cells = workbook.Worksheets(0).Cells

		Dim style As Aspose.Cells.Style

		'Get Style
		style = cells("B1").GetStyle()

		'Specify the fill color of the cell
		style.ForegroundColor = Color.Red
		style.Pattern = BackgroundType.Solid

		'Set Style
		cells("B1").SetStyle(style)

		'Get Style
		style = cells("B2").GetStyle()

		'Set the background, foreground colors of the cell
		style.ForegroundColor = Color.Yellow
		style.BackgroundColor = Color.Blue

		'Set Style Pattern
		style.Pattern = BackgroundType.DiagonalCrosshatch

		'Set Style
		cells("B2").SetStyle(style)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "PatternSetting.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "PatternSetting.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
