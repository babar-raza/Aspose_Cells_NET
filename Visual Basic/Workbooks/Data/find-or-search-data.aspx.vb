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

Partial Public Class Workbooks_Data_FindOrSearchData
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		'Call Method to create report
		CreateStaticReport()
	End Sub

	Protected Sub CreateStaticReport()
		'Instantiate a new workbook
		Dim workbook As New Workbook()

		'Set default font
		Dim style As Aspose.Cells.Style = workbook.DefaultStyle
		style.Font.Name = "Tahoma"
		workbook.DefaultStyle = style

		'Call Method to create data
		CreateSaticData(workbook)

		'Call Method to find data
		FindData(workbook)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "FindOrSearchData.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "FindOrSearchData.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub

	Private Shared Sub CreateSaticData(ByVal workbook As Workbook)
		'Get the cells collection in the first worksheet
		Dim cells As Cells = workbook.Worksheets(0).Cells

		'Put some values into cells
		cells("A1").PutValue("Product ID")
		cells("A2").PutValue(1)
		cells("A3").PutValue(2)
		cells("A4").PutValue(3)
		cells("A5").PutValue(4)

		cells("A7").PutValue(10)

		'Set a formula of the Cell. 
		cells("A7").Formula = "=SUM(A2:A5)"

		cells("B1").PutValue("Product Names")
		cells("B2").PutValue("Apples")
		cells("B3").PutValue("Bananas")
		cells("B4").PutValue("Grapes")
		cells("B5").PutValue("Oranges")
	End Sub

	Private Shared Sub FindData(ByVal workbook As Workbook)
		'Get the first worksheet
		Dim sheet As Worksheet = workbook.Worksheets(0)

		'Finds the cell with the input formula
		Dim cell1 As Aspose.Cells.Cell = sheet.Cells.FindFormula("=SUM(A2:A5)", Nothing)

		'Find the cell with formla which contains the input string
		Dim cell2 As Aspose.Cells.Cell = sheet.Cells.FindFormulaContains("SUM", Nothing)

		'Find the cell with the input integer or double
		Dim cell3 As Aspose.Cells.Cell = sheet.Cells.FindNumber(3, Nothing)

		'Find the cell with the input string
		Dim cell4 As Aspose.Cells.Cell = sheet.Cells.FindString("Apples", Nothing)

		'Find the cell containing with the input string
		Dim cell5 As Aspose.Cells.Cell = sheet.Cells.FindStringContains("anan", Nothing)

		'Find the cell ending with the input string
		Dim cell6 As Aspose.Cells.Cell = sheet.Cells.FindStringEndsWith("as", Nothing)

		'Find the cell starting with the input string
		Dim cell7 As Aspose.Cells.Cell = sheet.Cells.FindStringStartsWith("Gr", Nothing)

		Dim cells As Cells = workbook.Worksheets(0).Cells

		'Put some values into the cells
		cells("A9").PutValue("Name of the cell with the input formula (=SUM(A2:A5)): " & cell1.Name)
		cells("A10").PutValue("Name of the cell with formla which contains the input string (""SUM""): " & cell2.Name)
		cells("A11").PutValue("Name of the cell with the input integer or double (3): " & cell3.Name)
		cells("A12").PutValue("Name of the cell with the input string (""Apples""): " & cell4.Name)
		cells("A13").PutValue("Name of the cell containing with the input string (""anan""): " & cell5.Name)
		cells("A14").PutValue("Name of the cell ending with the input string (""as""): " & cell6.Name)
		cells("A15").PutValue("Name of the cell starting with the input string (""Gr""): " & cell7.Name)
	End Sub
End Class



