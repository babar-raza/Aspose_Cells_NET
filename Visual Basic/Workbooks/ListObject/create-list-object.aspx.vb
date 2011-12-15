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
Imports Aspose.Cells.Drawing
Imports Aspose.Cells.Tables


Partial Public Class CreateListObject
	Inherits System.Web.UI.Page
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub
	Public Shared Sub CreateStaticReport()

		'Open template from path
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
		path = path.Substring(0, path.LastIndexOf("\"))
		path &= "\designer\Workbooks\ListObject.xls"


		'Instantiate a new Workbook object.
		Dim workbook As New Workbook(path)

		'Get the first worksheet
		Dim sheet As Worksheet = workbook.Worksheets(0)

		'Get the ListObjects in the first sheet
		Dim listObjects As ListObjectCollection = sheet.ListObjects

		'Add a list object for the given data
		listObjects.Add(1, 1, 13, 5, True)

		'Set the totals visible
		listObjects(0).ShowTotals = True

		'Add the summary function to the last column in the list
		listObjects(0).ListColumns(4).TotalsCalculation = TotalsCalculation.Sum

		'Save the excel file
		workbook.Save(HttpContext.Current.Response, "ListObject.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))

		' End response to avoid unneeded html after xls
		HttpContext.Current.Response.End()


	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub
End Class
