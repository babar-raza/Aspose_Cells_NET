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

Partial Public Class Workbooks_DrawingObjects_AddingPictures
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Create Workbook
		Dim workbook As New Workbook()

		'Create worksheet
		Dim sheet As Worksheet = workbook.Worksheets(0)

		'Insert a picture into a cell
		Dim ImageUrl As String = System.Web.HttpContext.Current.Server.MapPath("~/Image/school.JPG")
		Dim pictureIndex As Integer = sheet.Pictures.Add(1, 1, ImageUrl)
		Dim picture As Aspose.Cells.Drawing.Picture = sheet.Pictures(pictureIndex)

		'Insert a picture into a cell using a stream
		Dim fs As FileStream = File.OpenRead(ImageUrl)
		'Create Byte Type array 
		Dim data(fs.Length - 1) As Byte
		'Read Data from stream into array
		fs.Read(data, 0, data.Length)
		'Close Stream
		fs.Close()

		'Crearte Memory Stream Object
		Dim stream As New MemoryStream()
		'Write data in memory
		stream.Write(data, 0, data.Length)

		'Create Image Object and load from stream
		Dim infoImage As System.Drawing.Image = System.Drawing.Image.FromStream(stream)
		'Insert a picture into a cell using a stream
		sheet.Pictures.Add(12, 1, stream, 100, 100)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "AddingPictures.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "AddingPictures.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
