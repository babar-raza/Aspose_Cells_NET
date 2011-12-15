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
Imports Aspose.Cells.Properties

Partial Public Class Workbooks_Worksheets_AddingDocumentProperties
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
		path &= "\designer\book1.xls"


		'Instantiate a Workbook object by calling its empty constructor
		Dim workbook As New Workbook(path)

		'Retrieve a list of all custom document properties of the Excel file
		Dim customProperties As CustomDocumentPropertyCollection = workbook.Worksheets.CustomDocumentProperties

		'Adding a custom document property to the Excel file
		Dim publisher As DocumentProperty = customProperties.Add("Publisher", "Aspose")

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "CustomDocumentProperties.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "CustomDocumentProperties.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub
End Class


