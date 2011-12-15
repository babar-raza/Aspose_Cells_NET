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

Partial Public Class Workbooks_Data_HelloWorld
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		'Call Method to create report
		CreateStaticReport()
	End Sub

	Protected Sub CreateStaticReport()
		'Open template
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
		path = path.Substring(0, path.LastIndexOf("\"))
		path &= "\designer\Workbooks\HelloWorld.xls"

		'Create a workbook object
		Dim workbook As New Workbook(path)


		'Get the first worksheet in the workbook
		Dim worksheet As Worksheet = workbook.Worksheets(0)

		'Get the cells collection in the sheet
		Dim cells As Cells = worksheet.Cells

		'Put a string value into the cell using its name
		cells("A1").PutValue("Cell Value")

		'put a string value into the cell using its name
		cells("A2").PutValue("Hello World")

		'Put an boolean value into the cell using its name
		cells("A3").PutValue(True)

		'Put an int value into the cell using its name
		cells("A4").PutValue(100)

		'Put an double value into the cell using its name
		cells("A5").PutValue(2856.5)

		'Put an string value that can be converted to other data type if appropriate
		cells("A6").PutValue((123.6).ToString(), True)

		'Put an object value into the cell using its name
		Dim obj As Object = "Aspose"
		cells("A7").PutValue(obj)

		'Put an datetime value into the cell
		Dim dt As DateTime = DateTime.Now
		cells("A8").PutValue(dt)
		Dim style As Aspose.Cells.Style = cells("A8").GetStyle()
		style.Font.IsBold = True
		cells("A8").SetStyle(style)

		'Put a string value into the cell using its row and column
		cells(0, 1).PutValue("Cell Value Type")

		For i As Integer = 1 To 7
			Select Case cells(i, 0).Type
				'Cell value is boolean
				Case CellValueType.IsBool
					cells(i, 1).PutValue("IsBool")
				'Cell value is datetime
				Case CellValueType.IsDateTime
					cells(i, 1).PutValue("IsDateTime")
				'Blank cell
				Case CellValueType.IsNull
					cells(i, 1).PutValue("IsNull")
				'Cell value is numeric
				Case CellValueType.IsNumeric
					cells(i, 1).PutValue("IsNumeric")
				'Cell value is string
				Case CellValueType.IsString
					cells(i, 1).PutValue("IsString")
				'Cell value type is unknown
				Case CellValueType.IsUnknown
					cells(i, 1).PutValue("IsUnknown")
			End Select
		Next i

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "HelloWorld.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "HelloWorld.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub

End Class
