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

Partial Public Class Workbooks_DrawingObjects_AddingImageHyperlink
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
		Dim worksheet As Worksheet = workbook.Worksheets(0)

		'Insert a picture into a cell
		Dim ImageUrl As String = System.Web.HttpContext.Current.Server.MapPath("~/Image/school.jpg")

		'Insert a string value to a cell
		worksheet.Cells("C2").PutValue("Image Hyperlink")

		'Set the 4th row height
		worksheet.Cells.SetRowHeight(3, 100)

		'Set the C column width
		worksheet.Cells.SetColumnWidth(2, 21)

		'Add a picture to the C4 cell
		Dim index As Integer = worksheet.Pictures.Add(3, 2, 4, 3, ImageUrl)

		'Get the picture object
		Dim pic As Aspose.Cells.Drawing.Picture = worksheet.Pictures(index)

		'Set the placement type
		pic.Placement = Aspose.Cells.Drawing.PlacementType.FreeFloating

		'Add an image hyperlink
		Dim hlink As Aspose.Cells.Hyperlink = pic.AddHyperlink("http://www.aspose.com/")

		'Specify the screen tip
		hlink.ScreenTip = "Click to go to Aspose site"

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "AddingImageHyperlink.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "AddingImageHyperlink.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
