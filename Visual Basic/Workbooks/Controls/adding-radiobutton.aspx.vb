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
Imports Aspose.Cells.Drawing

Partial Public Class Workbooks_Controls_AddRadioButton
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
		Dim excelbook As New Workbook()

		'Insert a value in C2 cell
		excelbook.Worksheets(0).Cells("C2").PutValue("Age Groups")

		'Get style from C2 cell
		Dim style As Aspose.Cells.Style = excelbook.Worksheets(0).Cells("C2").GetStyle()

		'Set the font text bold.
		style.Font.IsBold = True

		'Set style to C2 Cell
		excelbook.Worksheets(0).Cells("C2").SetStyle(style)

		'Add a radio button to the first sheet.
		Dim radio1 As Aspose.Cells.Drawing.RadioButton = excelbook.Worksheets(0).Shapes.AddRadioButton(3, 0, 2, 0, 30, 110)

		'Set its text string.
		radio1.Text = "20-29"

		'Set A1 cell as a linked cell for the radio button.
		radio1.LinkedCell = "A1"

		'Make the radio button 3-D.
		radio1.Shadow = True

		'Set the foreground color of the radio button.
		radio1.FillFormat.ForeColor = Color.LightGreen

		' set the line style of the radio button.
		radio1.LineFormat.Style = MsoLineStyle.ThickThin

		'Set the weight of the radio button.
		radio1.LineFormat.Weight = 4

		'Set the line color of the radio button.
		radio1.LineFormat.ForeColor = Color.Blue

		'Set the dash style of the radio button.
		radio1.LineFormat.DashStyle = MsoLineDashStyle.Solid

		'Make the line format visible.
		radio1.LineFormat.IsVisible = True

		'Make the fill format visible.
		radio1.FillFormat.IsVisible = True

		'Add another radio button to the first sheet.
		Dim radio2 As Aspose.Cells.Drawing.RadioButton = excelbook.Worksheets(0).Shapes.AddRadioButton(6, 0, 2, 0, 30, 110)

		'Set its text string.
		radio2.Text = "30-39"

		'Set A1 cell as a linked cell for the radio button.
		radio2.LinkedCell = "A1"

		'Make the radio button 3-D.
		radio2.Shadow = True

		'Set the foreground color of the radio button.
		radio2.FillFormat.ForeColor = Color.LightGreen

		' set the line style of the radio button.
		radio2.LineFormat.Style = MsoLineStyle.ThickThin

		'Set the weight of the radio button.
		radio2.LineFormat.Weight = 4

		'Set the line color of the radio button.
		radio2.LineFormat.ForeColor = Color.Blue

		'Set the dash style of the radio button.
		radio2.LineFormat.DashStyle = MsoLineDashStyle.Solid

		'Make the line format visible.
		radio2.LineFormat.IsVisible = True

		'Make the fill format visible.
		radio2.FillFormat.IsVisible = True

		'Add another radio button to the first sheet.
		Dim radio3 As Aspose.Cells.Drawing.RadioButton = excelbook.Worksheets(0).Shapes.AddRadioButton(9, 0, 2, 0, 30, 110)

		'Set its text string.
		radio3.Text = "40-49"

		'Set A1 cell as a linked cell for the radio button.
		radio3.LinkedCell = "A1"

		'Make the radio button 3-D.
		radio3.Shadow = True

		'Set the foreground color of the radio button.
		radio3.FillFormat.ForeColor = Color.LightGreen

		' set the line style of the radio button.
		radio3.LineFormat.Style = MsoLineStyle.ThickThin

		'Set the weight of the radio button.
		radio3.LineFormat.Weight = 4

		'Set the line color of the radio button.
		radio3.LineFormat.ForeColor = Color.Blue

		'Set the dash style of the radio button.
		radio3.LineFormat.DashStyle = MsoLineDashStyle.Solid

		'Make the line format visible.
		radio3.LineFormat.IsVisible = True

		'Make the fill format visible.
		radio3.FillFormat.IsVisible = True

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			excelbook.Save(HttpContext.Current.Response, "ComboBox.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			excelbook.Save(HttpContext.Current.Response, "ComboBox.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		' End response to avoid unneeded html after xls
		System.Web.HttpContext.Current.Response.End()
	End Sub

End Class



