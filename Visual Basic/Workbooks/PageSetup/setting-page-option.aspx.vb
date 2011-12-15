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

Partial Public Class Workbooks_PageSetup_SettingPageOption
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

		Dim workbook As New Workbook(path)

		Dim worksheet As Worksheet = workbook.Worksheets(0)
		'Set the orientation 
		worksheet.PageSetup.Orientation = PageOrientationType.Landscape

		'You can either choose FitToPages or Zoom property but not both at the same time
		worksheet.PageSetup.Zoom = 10
		'Set the paper size 
		worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4
		'Set the print quality of the worksheet 
		worksheet.PageSetup.PrintQuality = 200
		'Set the first page number of the worksheet pages
		worksheet.PageSetup.FirstPageNumber = 1

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "SettingPageOption.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "SettingPageOption.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
