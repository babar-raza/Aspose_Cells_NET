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
Imports Aspose.Cells.Drawing

Partial Public Class Workbooks_Controls_AddTextbox
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		'Call Method to create report
		CreateStaticReport()
	End Sub

	Protected Sub CreateStaticReport()
		'Instantiate a new Workbook.
		Dim workbook As New Workbook()

		'Get the first worksheet in the book.
		Dim worksheet As Worksheet = workbook.Worksheets(0)

		'Add a new textbox to the collection.
		Dim textboxIndex As Integer = worksheet.TextBoxes.Add(2, 1, 160, 200)

		'Get the textbox object.
		Dim textbox0 As Aspose.Cells.Drawing.TextBox = worksheet.TextBoxes(textboxIndex)

		'Fill the text.
		textbox0.Text = "ASPOSE______The .NET & JAVA Component Publisher!"

		'Get the textbox text frame.
		Dim textframe0 As Aspose.Cells.Drawing.MsoTextFrame = textbox0.TextFrame

		'Set the textbox to adjust it according to its contents.
		textframe0.AutoSize = True

		'Set the placement.
		textbox0.Placement = Aspose.Cells.Drawing.PlacementType.FreeFloating

		'Set the font color.
		textbox0.Font.Color = System.Drawing.Color.Blue

		'Set the font to bold.
		textbox0.Font.IsBold = True

		'Set the font size.
		textbox0.Font.Size = 14

		'Set font attribute to italic.
		textbox0.Font.IsItalic = True

		'Add a hyperlink to the textbox.
		textbox0.AddHyperlink("http://www.aspose.com/")

		'Get the filformat of the textbox.
		Dim fillformat As MsoFillFormat = textbox0.FillFormat

		'Set the fillcolor.
		fillformat.ForeColor = System.Drawing.Color.Silver

		'Get the lineformat type of the textbox.
		Dim lineformat As MsoLineFormat = textbox0.LineFormat

		'Set the line style.
		lineformat.Style = MsoLineStyle.ThinThick

		'Set the line weight.
		lineformat.Weight = 6

		'Set the dash style to squaredot.
		lineformat.DashStyle = MsoLineDashStyle.SquareDot

		'Add another textbox.
		textboxIndex = worksheet.TextBoxes.Add(15, 4, 85, 120)

		'Get the second textbox.
		Dim textbox1 As Aspose.Cells.Drawing.TextBox = worksheet.TextBoxes(textboxIndex)

		'Input some text to it.
		textbox1.Text = "This is another simple text box"

		'Set the placement type as the textbox will move and resize with cells.
		textbox1.Placement = Aspose.Cells.Drawing.PlacementType.MoveAndSize

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "TextBox.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "TextBox.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub

End Class



