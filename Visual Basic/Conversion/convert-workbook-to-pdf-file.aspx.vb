Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports System.Configuration
Imports System.Collections
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports Aspose.Cells

Partial Public Class Xls2Pdf
	Inherits System.Web.UI.Page
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)


	End Sub
	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Shared Sub CreateStaticReport()

		'Open template
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
		path = path.Substring(0, path.LastIndexOf("\"))
		path &= "\designer\MyTestBook1.xls"


		'Instantiate a new Workbook object.
		Dim book As New Workbook(path)

		'Save the workbook as a PDF File
		book.Save(HttpContext.Current.Response, "Xls2Pdf.pdf", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Pdf))

		'End response to avoid unneeded html after xls
		HttpContext.Current.Response.End()

	End Sub


End Class
