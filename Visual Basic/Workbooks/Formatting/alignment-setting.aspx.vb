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

Partial Public Class Workbooks_Formatting_AlignmentSetting
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Open template from path
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
		path = path.Substring(0, path.LastIndexOf("\"))
		path &= "\designer\Workbooks\AlignmentSetting.xls"


		'Create a new workbook
		Dim workbook As New Workbook(path)

		'Get the cells collection in the first worksheet
		Dim cells As Cells = workbook.Worksheets(0).Cells

		'Get Style Object 
		Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

		'Set text alignment type
		style.HorizontalAlignment = TextAlignmentType.Center
		style.VerticalAlignment = TextAlignmentType.Center

		'Set A1 style
		cells("A1").SetStyle(style)

		'Get Style Object 
		style = cells("A2").GetStyle()

		'Set text rotation angel
		style.RotationAngle = 45

		'Set A2 style
		cells("A2").SetStyle(style)

		'Get Style Object 
		style = cells("C3").GetStyle()

		'Set shrinktofit on
		style.ShrinkToFit = True

		'Set A3 style
		cells("C3").SetStyle(style)

		'Get Style Object 
		style = cells("A4").GetStyle()

		'Set the indentlevel
		style.IndentLevel = 5

		'Set A4 style
		cells("A4").SetStyle(style)

		'Get Style Object 
		style = cells("A5").GetStyle()

		'Wrapping Text
		style.IsTextWrapped = True

		'Set A5 style
		cells("A5").SetStyle(style)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "AlignmentSetting.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "AlignmentSetting.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class



