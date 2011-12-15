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

Partial Public Class Workbooks_Controls_AddCombobox
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		'Call Method to create report
		CreateStaticReport()
	End Sub

	Protected Sub CreateStaticReport()
		'Create a new Workbook.        
		Dim workbook As New Workbook()

		'Get the first worksheet.
		Dim sheet As Worksheet = workbook.Worksheets(0)

		'Get the worksheet cells collection.
		Dim cells As Cells = sheet.Cells

		'Input a value.
		cells("B3").PutValue("Employee:")
		Dim style As Aspose.Cells.Style = cells("B3").GetStyle()

		'Set it bold.
		style.Font.IsBold = True
		cells("B3").SetStyle(style)

		'Input some values that denote the input range for the combo box.
		cells("A2").PutValue("Emp001")

		cells("A3").PutValue("Emp002")

		cells("A4").PutValue("Emp003")

		cells("A5").PutValue("Emp004")

		cells("A6").PutValue("Emp005")

		cells("A7").PutValue("Emp006")

		'Add a new combo box.
		Dim comboBox As Aspose.Cells.Drawing.ComboBox = sheet.Shapes.AddComboBox(2, 0, 2, 0, 22, 100)

		'Set the linked cell;
		comboBox.LinkedCell = "A1"

		'Set the input range.
		comboBox.InputRange = "A2:A7"

		'Set no. of list lines displayed in the combo box's list portion.
		comboBox.DropDownLines = 5

		'Set the combo box with 3-D shading.
		comboBox.Shadow = True

		'AutoFit Columns
		sheet.AutoFitColumns()

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "ComboBox.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "ComboBox.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub

End Class



