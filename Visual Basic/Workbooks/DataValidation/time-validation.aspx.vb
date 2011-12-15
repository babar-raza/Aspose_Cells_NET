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

Partial Public Class TimeDataValidation
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		' Create a workbook.
		Dim workbook As New Workbook()

		' Obtain the cells of the first worksheet.
		Dim cells As Cells = workbook.Worksheets(0).Cells

		' Put a string value into A1 cell.
		cells("A1").PutValue("Please enter Time b/w 09:00 and 11:30 'o Clock")

		' Wrap the text.
		cells("A1").GetStyle().IsTextWrapped = True
		'cells["A1"].Style.IsTextWrapped = true;

		' Set the row height and column width for the cells.
		cells.SetRowHeight(0, 31)

		cells.SetColumnWidth(0, 35)

		' Get the validations collection.
		Dim validations As ValidationCollection = workbook.Worksheets(0).Validations

		' Add a new validation.
		Dim validation As Validation = validations(validations.Add())

		' Set the data validation type.
		validation.Type = ValidationType.Time

		' Set the operator for the data validation.
		validation.Operator = OperatorType.Between

		' Set the value or expression associated with the data validation.
		validation.Formula1 = "09:00"

		' The value or expression associated with the second part of the data validation.
		validation.Formula2 = "11:30"

		' Enable the error.
		validation.ShowError = True

		' Set the validation alert style.
		validation.AlertStyle = ValidationAlertType.Information

		' Set the title of the data-validation error dialog box.
		validation.ErrorTitle = "Time Error"

		' Set the data validation error message.
		validation.ErrorMessage = "Enter a Valid Time"

		' Set and enable the data validation input message.
		validation.InputMessage = "Time Validation Type"

		validation.IgnoreBlank = True

		validation.ShowInput = True

		' Set a collection of CellArea which contains the data validation settings.
		Dim cellArea As CellArea

		cellArea.StartRow = 0

		cellArea.EndRow = 0

		cellArea.StartColumn = 1

		cellArea.EndColumn = 1

		' Add the validation area.
		validation.AreaList.Add(cellArea)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "TimeValidation.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "TimeValidation.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
