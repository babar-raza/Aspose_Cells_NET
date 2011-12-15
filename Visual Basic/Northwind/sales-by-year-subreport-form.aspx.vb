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

Namespace Aspose.Cells.Demos.Northwind
	Partial Public Class SalesByYearSubreportForm
		Inherits System.Web.UI.Page
		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Define a workbook to store null initially
			Dim workbook As Workbook = Nothing

			Dim path As String = MapPath(".")
			path = path.Substring(0, path.LastIndexOf("\"))

			Dim salesByYearSubreport As New SalesByYearSubreport(path)

			'Create workbook based on the custom method of a class
			workbook = salesByYearSubreport.CreateSalesByYearSubreport()

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "SalesByYearSubreport.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "SalesByYearSubreport.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()

		End Sub
	End Class
End Namespace
