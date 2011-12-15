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

Partial Public Class Workbooks_Data_DataSorting
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
		path &= "\designer\Workbooks\unsorted.xls"


		'Instantiate a new Workbook object.
		Dim workbook As New Workbook(path)

		'Get the workbook datasorter object.
		Dim sorter As DataSorter = workbook.DataSorter

		'Set the first order for datasorter object.
		sorter.Order1 = Aspose.Cells.SortOrder.Descending

		'Define the first key.
		sorter.Key1 = 0

		'Set the second order for datasorter object.
		sorter.Order2 = Aspose.Cells.SortOrder.Ascending

		'Define the second key.
		sorter.Key2 = 1

		'Create a cells area (range).
		Dim ca As New CellArea()

		'Specify the start row index.
		ca.StartRow = 0

		'Specify the start column index.
		ca.StartColumn = 0

		'Specify the last row index.
		ca.EndRow = 13

		'Specify the last column index.
		ca.EndColumn = 1

		'Sort data in the specified data range (A1:B14)
		sorter.Sort(workbook.Worksheets(0).Cells, ca)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "DataSorting.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "DataSorting.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub

End Class



