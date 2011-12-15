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

Partial Public Class Workbooks_DrawingObjects_OtherDrawingObjects
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

		'Create Worksheet
		Dim sheet As Worksheet = workbook.Worksheets(0)

		'Add Textbox object in collection
		Dim textboxIndex As Integer = sheet.TextBoxes.Add(1, 1, 40, 40)

		'Get newly added Textbox from collection
		Dim textbox As Aspose.Cells.Drawing.TextBox = sheet.TextBoxes(textboxIndex)

		'Set TextBox Text
		textbox.Text = "Sample Text Box"

		'Set Textbox dimensions
		textbox.Height = 80
		textbox.Width = 80

		'Get path of Image in Variable
		Dim imageUrl As String = System.Web.HttpContext.Current.Server.MapPath("~/Image/school.jpg")

		'Create File Stream to read image Data
		Dim fs As FileStream = File.OpenRead(imageUrl)

		'Initialize Byte Array to store Image Data
		Dim imageData(fs.Length - 1) As Byte

		'Read File Stream Data into Array
		fs.Read(imageData, 0, imageData.Length)

		'Cloese File Stream
		fs.Close()

		'Open template
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
		path = path.Substring(0, path.LastIndexOf("\"))
		path &= "\designer\book1.xls"

		'Read Template file through Stream
		fs = File.OpenRead(path)

		'Create Byte array to store Template file data
		Dim objectData(fs.Length - 1) As Byte

		'Start read Data
		fs.Read(objectData, 0, objectData.Length)

		'Close File Stream
		fs.Close()

		'Add Image as Ole Objects to Worksheet OleObjects Collection
		sheet.OleObjects.Add(3, 3, 150, 150, imageData)

		' embedded ole object data 
		sheet.OleObjects(0).ObjectData = objectData

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "OtherDrawingObjects.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "OtherDrawingObjects.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub
End Class
