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

Partial Public Class Workbooks_Formatting_FontSetting
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Create a workbook
		Dim workbook As New Workbook()

		'Get the cells collection in the first worksheet
		Dim cells As Cells = workbook.Worksheets(0).Cells

		'Put a value into the cell
		cells("A2").PutValue("Aspose")

		'Get Style
		Dim style As Aspose.Cells.Style = cells("A2").GetStyle()

		'Set the color of the font
		style.Font.Color = Color.Red

		'Set Style
		cells("A2").SetStyle(style)

		'Put a value into the cell
		cells("B2").PutValue("Aspose")

		'Get Style
		style = cells("B2").GetStyle()

		'Set a value indicating whether the font is bold
		style.Font.IsBold = True

		'Set Style
		cells("B2").SetStyle(style)

		'Put a value into the cell
		cells("C2").PutValue("Aspose")

		'Get Style
		style = cells("C2").GetStyle()

		'Set a value indicating whether the font is italic
		style.Font.IsItalic = True

		'Set Style
		cells("C2").SetStyle(style)

		'Put a value into the cell
		cells("A4").PutValue("Aspose")

		'Get Style
		style = cells("A4").GetStyle()

		'Set a value indicating whether the font is strikeout
		style.Font.IsStrikeout = True

		'Set Style
		cells("A4").SetStyle(style)

		'Put a value into the cell
		cells("B4").PutValue("Aspose")

		'Get Style
		style = cells("B4").GetStyle()

		'Set a value indicating whether the font is subscript
		style.Font.IsSubscript = True

		'Set Style
		cells("B4").SetStyle(style)

		'Put a value into the cell
		cells("C4").PutValue("Aspose")

		'Get Style
		style = cells("C4").GetStyle()

		'Set a value indicating whether the font is super script.
		style.Font.IsSuperscript = True

		'Set Style
		cells("C4").SetStyle(style)

		'Put a value into the cell
		cells("A6").PutValue("Aspose")

		'Get Style
		style = cells("A6").GetStyle()

		'Set the name of the font
		style.Font.Name = "Verdana"

		'Set Style
		cells("A6").SetStyle(style)

		'Put a value into the cell
		cells("B6").PutValue("Aspose")

		'Get Style
		style = cells("B6").GetStyle()

		'Set the size of the font
		style.Font.Size = 15

		'Set Style
		cells("B6").SetStyle(style)

		'Put a value into the cell
		cells("C6").PutValue("Aspose")

		'Get Style
		style = cells("C6").GetStyle()

		'Set the font underline type
		style.Font.Underline = FontUnderlineType.Accounting

		'Set Style
		cells("C6").SetStyle(style)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "FontSetting.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "FontSetting.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class




