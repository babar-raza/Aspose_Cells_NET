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

Partial Public Class Workbooks_Formatting_ModifyExistingStyle
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Create a workbook.
		Dim workbook As New Workbook()

		'Create a new style object.
		Dim style As Aspose.Cells.Style = workbook.Styles(workbook.Styles.Add())

		'Set the number format.
		style.Number = 14

		'Set the font color to red color.
		style.Font.Color = System.Drawing.Color.Red

		'Name the style.
		style.Name = "Style1"

		'Get the first worksheet cells.
		Dim cells As Cells = workbook.Worksheets(0).Cells

		'Put value in cell
		cells("A1").PutValue("Original Color Red & Modified Color Blue")

		'Specify the style (described above) to A1 cell.
		cells("A1").SetStyle(style)

		'Create a range (B1:D1).
		Dim range As Range = cells.CreateRange("B1", "D1")

		'Initialize styleflag struct.
		Dim flag As New StyleFlag()

		'Set all formatting attributes on.
		flag.All = True

		'Apply the style (described above)to the range.
		range.ApplyStyle(style, flag)

		'Modify the style (described above) and change the font color from red to blue.
		style.Font.Color = System.Drawing.Color.Blue

		'Done! Since the named style (described above) has been set to a cell and range, 
		'the change would be Reflected(new modification is implemented) to cell(A1) and //range (B1:D1).
		style.Update()

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "ModifyExistingStyle.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "ModifyExistingStyle.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub

End Class



