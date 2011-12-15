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

Partial Public Class DateDataValidation
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

		Dim dateFirst As String = DateTime.Now.AddDays(-10).ToShortDateString()
		Dim dateSecond As String = DateTime.Now.ToShortDateString()

		Dim message As String = "Enter date btw """ & dateFirst & """ and """ & dateSecond & """ in cell A2."

		' Put a string value into the A1 cell.
		cells("A1").PutValue(message)

		' Set row height and column width for the cells.
		cells.SetRowHeight(0, 31)

		cells.SetColumnWidth(0, 35)

		' Get the validations collection.
		Dim validations As ValidationCollection = workbook.Worksheets(0).Validations

		' Add a new validation.
		Dim validation As Validation = validations(validations.Add())

		' Set the data validation type.
		validation.Type = ValidationType.Date

		' Set the operator for the data validation
		validation.Operator = OperatorType.Between

		' Set the value or expression associated with the data validation.
		validation.Formula1 = dateFirst

		' The value or expression associated with the second part of the data validation.
		validation.Formula2 = dateSecond

		' Enable the error.
		validation.ShowError = True

		' Set the validation alert style.
		validation.AlertStyle = ValidationAlertType.Stop

		' Set the title of the data-validation error dialog box
		validation.ErrorTitle = "Date Error"

		' Set the data validation error message.
		validation.ErrorMessage = "Invalid Date. " & message

		' Set and enable the data validation input message.
		validation.InputMessage = "Date Validation Type"

		validation.IgnoreBlank = True

		validation.ShowInput = True

		' Set a collection of CellArea which contains the data validation settings.
		Dim cellArea As CellArea

		cellArea.StartRow = 1

		cellArea.EndRow = 1

		cellArea.StartColumn = 0

		cellArea.EndColumn = 0

		' Add the validation area.
		validation.AreaList.Add(cellArea)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "DateValidation.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "DateValidation.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
