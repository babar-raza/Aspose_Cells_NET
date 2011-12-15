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

Partial Public Class Workbooks_Formatting_RichTextFormatting
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		 'Instantiating an Workbook object
		Dim workbook As New Workbook()

		'Obtaining the reference of the newly added worksheet by passing its sheet index
		Dim worksheet As Worksheet = workbook.Worksheets(0)

		'Accessing the "A1" cell from the worksheet
		Dim cell As Aspose.Cells.Cell = worksheet.Cells("A1")

		'Adding some value to the "A1" cell
		cell.PutValue("Rich Text Formatting Demo")

		cell.Characters(3, 15).Font.IsItalic = True

		cell.Characters(5, 4).Font.Name = "Algerian"

		cell.Characters(21, 4).Font.Color = System.Drawing.Color.Red

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "RichTextFormatting.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "RichTextFormatting.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub

End Class



