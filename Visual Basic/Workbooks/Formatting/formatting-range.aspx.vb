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

Partial Public Class Workbooks_Formatting_FormattingRange
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Create a new workbook
		Dim workbook As New Workbook()
		'Get the first worksheet
		Dim sheet As Worksheet = workbook.Worksheets(0)
		'Get its cells collection
		Dim cells As Cells = sheet.Cells

		'Create a named range
		Dim range As Range = sheet.Cells.CreateRange("B1", "E5")
		'Set the name of the named range
		range.Name = "Range1"
		'Create a new style adding to the workbook styles collection
		Dim style As Aspose.Cells.Style = workbook.Styles(workbook.Styles.Add())
		'Specify the style's fill color
		style.ForegroundColor = System.Drawing.Color.Blue
		style.Pattern = BackgroundType.Solid

		'Create a styleflag object
		Dim styleFlag As New StyleFlag()
		'Specify all attributes
		styleFlag.All = True
		'Apply the style to the range
		range.ApplyStyle(style, styleFlag)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "FormattingRange.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "FormattingRange.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub

End Class



