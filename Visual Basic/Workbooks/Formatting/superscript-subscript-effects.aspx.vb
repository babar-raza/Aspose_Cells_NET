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

Partial Public Class Workbooks_Formatting_SuperscriptSubscript
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Instantiating a Workbook object
		Dim workbook As New Workbook()

		'Obtaining the reference to the first worksheet
		Dim worksheet As Worksheet = workbook.Worksheets(0)

		'Accessing the "A1" cell from the worksheet
		Dim cell As Cell = worksheet.Cells("A1")

		'Get Style
		Dim style As Aspose.Cells.Style = cell.GetStyle()

		'Adding some value to the "A1" cell
		cell.PutValue("Hello")

		'Setting the font Superscript
		style.Font.IsSuperscript = True

		'Set Style
		cell.SetStyle(style)

		'Get Cell
		cell = worksheet.Cells("A2")

		'Get Style
		style = cell.GetStyle()

		'Adding some value to the "A2" cell
		cell.PutValue("Aspose")

		'Setting the font Superscript
		style.Font.IsSubscript = True

		'Set Style
		cell.SetStyle(style)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "SuperscriptSubscript.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "SuperscriptSubscript.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub

End Class



