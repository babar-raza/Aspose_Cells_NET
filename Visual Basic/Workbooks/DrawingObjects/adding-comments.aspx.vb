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

Partial Public Class Workbooks_DrawingObjects_AddingComments
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Create Workbook
		Dim workbook As New Workbook()

		'Create Worksheet
		Dim worksheet As Worksheet = workbook.Worksheets(0)

		'Create Cells
		Dim cells As Cells = worksheet.Cells

		'Put a value into a cell
		cells("B1").PutValue("Hello")

		'Add comment to cell B1
		Dim commentIndex As Integer = worksheet.Comments.Add(0, 1)

		'Access the newly added comment
		Dim comment As Comment = worksheet.Comments(commentIndex)

		'Set the comment note
		comment.Note = "Aspose.Cells"

		'Set the font of a comment
		comment.Font.Size = 12
		comment.Font.IsBold = True
		comment.HeightCM = 5
		comment.WidthCM = 5

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "AddingComments.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "AddingComments.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
