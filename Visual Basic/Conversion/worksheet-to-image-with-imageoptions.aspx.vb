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
Imports Aspose.Cells.Rendering

Partial Public Class Sheet2ImageWithOptions
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

		'Get the first worksheet
		Dim sheet As Worksheet = book.Worksheets(0)

		'Apply different Image and Print options 
		Dim options As New ImageOrPrintOptions()
		options.HorizontalResolution = 300
		options.VerticalResolution = 300
		options.TiffCompression = TiffCompression.CompressionCCITT4
		options.IsCellAutoFit = False
		options.ImageFormat = System.Drawing.Imaging.ImageFormat.Tiff
		options.PrintingPage = PrintingPageType.Default

		'Create a memory stream object.
		Dim memorystream As New MemoryStream()

		Dim sheetRender As New SheetRender(sheet, options)

		'Convert worksheet to image.
		sheetRender.ToTiff(memorystream)

		memorystream.Seek(0, SeekOrigin.Begin)

		'Set Response object to stream the image file.
		Dim data() As Byte = memorystream.ToArray()
		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.ContentType = "image/tiff"
		HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=SheetImage.tiff")
		HttpContext.Current.Response.OutputStream.Write(data, 0, data.Length)

		'End response to avoid unneeded html after xls
		HttpContext.Current.Response.End()

	End Sub


End Class
