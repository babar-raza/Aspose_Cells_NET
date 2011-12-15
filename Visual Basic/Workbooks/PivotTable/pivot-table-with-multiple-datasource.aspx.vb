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
Imports Aspose.Cells.Pivot

Namespace Aspose.Cells.Demos
	Partial Public Class Pivot_Table_MultiSource
		Inherits System.Web.UI.Page
		Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
			CreateStaticReport()
		End Sub

		Public Sub CreateStaticReport()
			'Open template
			Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
			path = path.Substring(0, path.LastIndexOf("\"))

			path &= "\designer\Workbooks\PivotSource.xls"
			'Instantiating an Workbook object
			Dim workbook As New Workbook(path)

			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim pivotTables As PivotTableCollection = sheet.PivotTables

			Dim sourceData() As String = { "=Sheet1!A1:C8", "=Sheet2!A1:C8" }
			Dim pageField As New PivotPageFields()
			Dim pageItems(1) As String
			pageItems(0) = "Item1"
			pageItems(1) = "Item2"
			pageField.AddPageField(pageItems)
			pageItems = New String(1){}
			pageItems(0) = "Item3"
			pageItems(1) = "Item4"
			pageField.AddPageField(pageItems)
			Dim TBPG(1) As Integer

			TBPG(0) = 0
			TBPG(1) = 1

			'Sets which item label in each page field to use to identify the data range.
			pageField.AddIdentify(0, TBPG)
			TBPG = New Integer(1){}
			TBPG(0) = 1
			TBPG(1) = -1
			pageField.AddIdentify(1, TBPG)
			Dim index As Integer = pivotTables.Add(sourceData, False, pageField, "E3", "PivotTable1")

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "PivotTableMultipleSource.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "PivotTableMultipleSource.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub
	End Class
End Namespace