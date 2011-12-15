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
Imports System.Drawing

Partial Public Class Workbooks_Data_AddingHyperlinks
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		'Call Method to create report
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Create a new Workbook.        
		Dim workbook As New Workbook()

		'Get the first worksheet.
		Dim worksheet As Worksheet = workbook.Worksheets(0)

		'Get cells from workbook
		Dim cells As Cells = worksheet.Cells

		'Put a value into a cell
		cells("A1").PutValue("Visit Aspose")

		'Get Style Object 
		Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

		'Set the font color of the cell to Blue
		style.Font.Color = Color.Blue

		'Set the font of the cell to Single Underline
		style.Font.Underline = FontUnderlineType.Single

		'Set the style of A1 cell
		cells("A1").SetStyle(style)

		'Add a hyperlink to Aspose web sit at cell "A1"
		worksheet.Hyperlinks.Add("A1", 1, 1, "http://www.aspose.com")

		'add a hyperlink to another cell at cell "C1"
		worksheet.Hyperlinks.Add("C1", 1, 1, "Sheet1!A10")

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "AddingHyperlinks.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "AddingHyperlinks.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub
End Class
