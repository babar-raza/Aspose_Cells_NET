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

Partial Public Class Workbooks_Data_CalculateFormula
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		'Call Method to create report
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Open template
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
		path = path.Substring(0, path.LastIndexOf("\"))
		path &= "\designer\Workbooks\CalculateFormula.xls"

		'Instantiate a workbook
		Dim workbook As New Workbook(path)

		'Get the cells collection in the first worksheet
		Dim cells As Cells = workbook.Worksheets(0).Cells
		For i As Integer = 11 To 85
			'Get a string value from a cell
			Dim strFormula As String = cells(i, 2).StringValue
			'Set a formula of the Cell
			cells(i, 3).Formula = strFormula
		Next i

		'Calculates the result of formulas
		workbook.CalculateFormula()
		For i As Integer = 11 To 85
			'Put values obtaining the calculated values
			cells(i, 4).PutValue(cells(i, 3).Value)
		Next i

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "CalculateFormula.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "CalculateFormula.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub
End Class

