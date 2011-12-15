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

Partial Public Class ListDataValidation
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		' Create a workbook object.
		Dim workbook As New Workbook()

		' Get the first worksheet.
		Dim worksheet1 As Worksheet = workbook.Worksheets(0)

		' Add a new worksheet and access it.
		Dim i As Integer = workbook.Worksheets.Add()

		Dim worksheet2 As Worksheet = workbook.Worksheets(i)

		' Create a range in the second worksheet.
		Dim range As Range = worksheet2.Cells.CreateRange("E1", "E4")

		' Name the range.
		range.Name = "MyRange"

		' Fill different cells with data in the range.
		range(0, 0).PutValue("Blue")
		range(1, 0).PutValue("Red")
		range(2, 0).PutValue("Green")
		range(3, 0).PutValue("Yellow")

		' Get the validations collection.
		Dim validations As ValidationCollection = worksheet1.Validations

		' Create a new validation to the validations list.
		Dim validation As Validation = validations(validations.Add())

		' Set the validation type.
		validation.Type = Aspose.Cells.ValidationType.List

		' Set the operator.
		validation.Operator = OperatorType.None

		' Set the in cell drop down.
		validation.InCellDropDown = True

		' Set the formula1.
		validation.Formula1 = "=MyRange"

		' Enable it to show error.
		validation.ShowError = True

		' Set the alert type severity level.
		validation.AlertStyle = ValidationAlertType.Stop

		' Set the error title.
		validation.ErrorTitle = "Error"

		' Set the error message.
		validation.ErrorMessage = "Please select a color from the list"

		' Specify the validation area.
		Dim area As CellArea
		area.StartRow = 0
		area.EndRow = 4
		area.StartColumn = 0
		area.EndColumn = 0

		' Add the validation area.
		validation.AreaList.Add(area)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "ListDataValidation.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "ListDataValidation.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
