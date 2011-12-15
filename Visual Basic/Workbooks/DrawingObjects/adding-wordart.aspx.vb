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

Partial Public Class Workbooks_DrawingObjects_AddingWordArt
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'initialize the workbook object
		Dim workbook As New Workbook()

		'Get first worksheet in the workbook
		Dim sheet As Worksheet = workbook.Worksheets(0)

		'Apply WordArt Style with font settings
		sheet.Shapes.AddTextEffect(Aspose.Cells.Drawing.MsoPresetTextEffect.TextEffect1, "Aspose.Cells for .NET", "Arial", 15, True, True, 5, 5, 2, 2, 100, 175)

		'Apply WordArt Style with font settings
		sheet.Shapes.AddTextEffect(Aspose.Cells.Drawing.MsoPresetTextEffect.TextEffect2, "Aspose.Cells for Java", "Verdana", 30, True, False, 10, 5, 2, 2, 100, 100)

		'Apply WordArt Style with font settings
		sheet.Shapes.AddTextEffect(Aspose.Cells.Drawing.MsoPresetTextEffect.TextEffect3, "Aspose.Cells for Reporting Services", "Times New Roman", 25, False, True, 15, 5, 2, 2, 100, 150)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "AddingWordArt.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "AddingWordArt.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
