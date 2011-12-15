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
Imports Aspose.Cells.Drawing

Partial Public Class Workbooks_DrawingObjects_ExtractingOleObject
	Inherits System.Web.UI.Page
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Shared Sub CreateStaticReport()
		'Open template from path
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
		path = path.Substring(0, path.LastIndexOf("\"))
		path &= "\designer\OleFile.xls"

		'Instantiating a Workbook object
		Dim workbook As New Workbook(path)

		'Get the OleObject Collection in the first worksheet.
		Dim oles As OleObjectCollection = workbook.Worksheets(0).OleObjects

		'Loop through all the oleobjects and extract each object in the worksheet.
		For i As Integer = 0 To oles.Count - 1
			'Create Ole Object and Initialize it with i Item in collection
			Dim ole As OleObject = oles(i)

			'Specify the output filename.
			Dim fileName As String = "outOle" & i & "."

			'Specify each file format based on the oleobject format type.
			Select Case ole.FileType

				Case OleFileType.Doc
					fileName &= "doc"

				Case OleFileType.Xls
					fileName &= "Xls"

				Case OleFileType.Ppt
					fileName &= "Ppt"

				Case OleFileType.Pdf
					fileName &= "Pdf"

				Case OleFileType.Unknown
					fileName &= "Jpg"

				Case Else
					'........
			End Select


			'Save the oleobject as a new excel file if the object type is xls.
			If ole.FileType = OleFileType.Xls Then
				'Create MemoryStream
				Dim ms As New MemoryStream()

				'Write OleObject to Memory Stream 
				ms.Write(ole.ObjectData, 0, ole.ObjectData.Length)

				'Ctreate WorkBook from MemoryStream
				Dim oleBook As New Workbook(ms)

				'Hide all worksheets from workbook
				oleBook.Worksheets.IsHidden = False

				'Saving the Excel file
				oleBook.Save(HttpContext.Current.Response,"OleObect.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))


		   'Create the files based on the oleobject format types.                
			Else

				'FileStream fs = File.Create(fileName);
				HttpContext.Current.Response.Clear()
				HttpContext.Current.Response.ContentType = "image/jpg"
				HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=OleFile.jpg")
				HttpContext.Current.Response.OutputStream.Write(ole.ObjectData, 0, ole.ObjectData.Length)

			End If

		Next i
		 ' End response to avoid unneeded html after xls
		HttpContext.Current.Response.End()
	End Sub
End Class
