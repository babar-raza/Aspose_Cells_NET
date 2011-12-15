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

Partial Public Class WholeNumberValidation
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Initialize Workbook
		Dim workbook As New Workbook()

		'Initialize WorkSheet
		Dim worksheet As Worksheet = workbook.Worksheets(0)

		' Obtain the cells of the first worksheet.
		Dim cells As Cells = workbook.Worksheets(0).Cells

		'Put a string value into A1 cell.
		cells("A1").PutValue("Please enter whole number between 10 and 1000 only in this column.")

		'Accessing the Validations collection of the worksheet
		Dim validations As ValidationCollection = worksheet.Validations

		'Creating a Validation object
		Dim validation As Validation = validations(validations.Add())

		'Setting the validation type to whole number
		validation.Type = ValidationType.WholeNumber

		'Setting the operator for validation to Between
		validation.Operator = OperatorType.Between

		'Setting the minimum value for the validation
		validation.Formula1 = "10"

		'Setting the maximum value for the validation
		validation.Formula2 = "1000"

		validation.ErrorMessage = "Invalid Whole Number. Enter whole number between 10 and 1000 only."

		'Applying the validation to a range of cells from A1 to B2 using the CellArea structure
		Dim area As CellArea
		area.StartRow = 1
		area.EndRow = 9
		area.StartColumn = 0
		area.EndColumn = 0

		'Adding the cell area to Validation
		validation.AreaList.Add(area)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "WholeNumberValidation.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "WholeNumberValidation.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
