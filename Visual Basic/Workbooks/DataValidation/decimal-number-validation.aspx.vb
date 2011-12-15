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

Partial Public Class DecimalNumberValidation
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

		' Create a worksheet and get the first worksheet.
		Dim ExcelWorkSheet As Worksheet = workbook.Worksheets(0)

		' Obtain the existing Validations collection.
		Dim validations As ValidationCollection = ExcelWorkSheet.Validations

		' Create a validation object adding to the collection list.
		Dim validation As Validation = validations(validations.Add())

		' Set the validation type.
		validation.Type = ValidationType.Decimal

		' Specify the operator.
		validation.Operator = OperatorType.Between

		' Set the lower and upper limits.
		validation.Formula1 = Decimal.MinValue.ToString()

		validation.Formula2 = Decimal.MaxValue.ToString()

		' Set the error message.
		validation.ErrorMessage = "Please enter a valid integer or decimal number"

		' Specify the validation area of cells.
		Dim area As CellArea
		area.StartRow = 0
		area.EndRow = 9
		area.StartColumn = 0
		area.EndColumn = 0

		' Add the area.
		validation.AreaList.Add(area)

		' Set the number formats to 2 decimal places for the validation area.        

		For i As Integer = 0 To 9
			Dim style As New Aspose.Cells.Style()
			style.Custom = "0.00"
		   ExcelWorkSheet.Cells(i, 0).SetStyle(style)
		Next i

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "DecimalNumberValidation.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "DecimalNumberValidation.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
