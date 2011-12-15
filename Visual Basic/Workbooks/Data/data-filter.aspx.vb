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

Partial Public Class Workbooks_Data_DataFilter
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)

		'Call Method to create report
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Create a new workbook
		Dim workbook As New Workbook()

		'Get the first worksheet in the workbook
		Dim sheet As Worksheet = workbook.Worksheets(0)

		'Get the cells collection in the sheet
		Dim cells As Cells = sheet.Cells

		'Put some values into cells 
		cells("A1").PutValue("Fruit")
		cells("B1").PutValue("Total")
		cells("A2").PutValue("Apple")
		cells("B2").PutValue(1000)
		cells("A3").PutValue("Orange")
		cells("B3").PutValue(2500)
		cells("A4").PutValue("Bananas")
		cells("B4").PutValue(2500)
		cells("A5").PutValue("Pear")
		cells("B5").PutValue(1000)
		cells("A6").PutValue("Grape")
		cells("B6").PutValue(2000)

		cells("D1").PutValue("Count:")

		'Set a formula to E1 cell
		cells("E1").Formula = "=SUBTOTAL(2,B1:B6)"

		'Represents the range to which the specified AutoFilter applies
		sheet.AutoFilter.Range = "A1:B6"

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "DataFilteringAndValidation.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "DataFilteringAndValidation.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub

End Class



