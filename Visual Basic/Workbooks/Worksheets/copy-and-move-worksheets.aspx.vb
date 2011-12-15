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

Partial Public Class Copy_Move
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
		path &= "\designer\Workbooks\Copy_Move.xls"

		'Instantiate a new Workbook object.
		Dim workbook As New Workbook(path)

		'Copy the first sheet contents into the last worksheet in the book
		workbook.Worksheets(2).Copy(workbook.Worksheets("Copy"))

		'Move the sheet to the last indexed position in the book 
		workbook.Worksheets("Move").Move(2)

		'Save the excel file
		workbook.Save(HttpContext.Current.Response, "CopyandMove.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))

		' End response to avoid unneeded html after xls
		HttpContext.Current.Response.End()
	End Sub
End Class
