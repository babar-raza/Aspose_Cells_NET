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

Partial Public Class Workbooks_Controls_AddCheckbox
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		'Call Method to create report
		CreateStaticReport()
	End Sub

	Protected Sub CreateStaticReport()
		'Instantiate a new Workbook.
		Dim workbook As New Workbook()

		'Add a checkbox to the first worksheet in the workbook.
		Dim index As Integer = workbook.Worksheets(0).CheckBoxes.Add(5, 5, 20, 120)

		'Get the checkbox object.
		Dim checkbox As Aspose.Cells.Drawing.CheckBox = workbook.Worksheets(0).CheckBoxes(index)

		'Set its text string.
		checkbox.Text = "Click it!"

		'Put a value into B1 cell.
		workbook.Worksheets(0).Cells("B1").PutValue("LnkCell")

		'Set B1 cell as a linked cell for the checkbox.
		checkbox.LinkedCell = "B1"

		'Check the checkbox by default.
		checkbox.Value = True

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "CheckBox.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "CheckBox.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub

End Class



