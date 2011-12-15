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
Imports System.IO

Partial Public Class Workbooks_DrawingObjects_AddingImageFromWeb
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Define memory stream object
		Dim objImage As System.IO.MemoryStream

		'Define web client object
		Dim objwebClient As System.Net.WebClient

		'Define a string which will hold the web image url
		Dim sURL As String = "http://www.xlsoft.com/jp/products/aspose/images/Aspose_Cells-Product-Box.jpg"

		'Instantiate the web client object
		objwebClient = New System.Net.WebClient()

		'Now, extract data into memory stream downloading the image data into the array of bytes
		objImage = New System.IO.MemoryStream(objwebClient.DownloadData(sURL))

		'Create a new workbook
		Dim workbook As New Aspose.Cells.Workbook()

		'Get the first worksheet in the book
		Dim sheet As Aspose.Cells.Worksheet = workbook.Worksheets(0)

		'Get the first worksheet pictures collection
		Dim pictures As Aspose.Cells.Drawing.PictureCollection = sheet.Pictures

		'Insert the picture from the stream to B2 cell
		pictures.Add(1, 1, objImage)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "AddingImageFromWeb.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "AddingImageFromWeb.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
