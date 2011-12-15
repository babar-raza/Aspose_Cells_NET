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
Imports System.IO
Imports Aspose.Cells

Partial Public Class Workbooks_DrawingObjects_InsertingOleObject
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Instantiate a new Workbook.
		Dim workbook As New Workbook()

		'Get the first worksheet. 
		Dim sheet As Worksheet = workbook.Worksheets(0)

		'Define a string variable to store the image path.
		Dim ImageUrl As String = System.Web.HttpContext.Current.Server.MapPath("~/Image/school.JPG")

		'Get the picture into the streams.
		Dim fs As FileStream = File.OpenRead(ImageUrl)

		'Define a byte array.
		Dim imageData(fs.Length - 1) As Byte

		'Obtain the picture into the array of bytes from streams.
		fs.Read(imageData, 0, imageData.Length)

		'Close the stream.
		fs.Close()

		'Get an excel file path in a variable.
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~/designer/book1.xls")

		'Get the file into the streams.
		fs = File.OpenRead(path)

		'Define an array of bytes. 
		Dim objectData(fs.Length - 1) As Byte

		'Store the file from streams.
		fs.Read(objectData, 0, objectData.Length)

		'Close the stream.
		fs.Close()

		'Add an Ole object into the worksheet with the image
		'shown in MS Excel.
		sheet.OleObjects.Add(4, 3, 200, 200, imageData)

		'Set embedded ole object data.     
		sheet.OleObjects(0).ObjectData = objectData

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "InsertOleObect.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "InsertOleObect.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
