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

Partial Public Class Workbooks_Formatting_ConditionalFormatting
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Instantiating a Workbook object
		Dim workbook As New Workbook()

		Dim sheet As Worksheet = workbook.Worksheets(0)

		'Adds an empty conditional formatting
		Dim index As Integer = sheet.ConditionalFormattings.Add()

		'Initialize FormatConditionCollection from newly inserted Index
		Dim fcs As FormatConditionCollection = sheet.ConditionalFormattings(index)

		'Sets the conditional format range.
		Dim ca As New CellArea()
		ca.StartRow = 0
		ca.EndRow = 0
		ca.StartColumn = 0
		ca.EndColumn = 0

		'Assign FormatConditionCollection the Area
		fcs.AddArea(ca)


		'Adds condition.
		Dim conditionIndex As Integer = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "50", "100")

		'Sets the background color.
		Dim fc As FormatCondition = fcs(conditionIndex)

		'Set BackgroundColor
		fc.Style.BackgroundColor = Color.Red



		'Adds an empty conditional formatting
		Dim index2 As Integer = sheet.ConditionalFormattings.Add()

		'Initialize FormatConditionCollection for newly added index
		Dim fcs2 As FormatConditionCollection = sheet.ConditionalFormattings(index2)

		'Sets the conditional format range.
		Dim ca2 As New CellArea()
		ca2.StartRow = 2
		ca2.EndRow = 2
		ca2.StartColumn = 1
		ca2.EndColumn = 1

		'Assign FormatConditionCollection the Area
		fcs2.AddArea(ca2)


		'Adds condition.
        Dim conditionIndex2 As Integer = fcs2.AddCondition(FormatConditionType.Expression)

		'Sets the background color.
		Dim fc2 As FormatCondition = fcs2(conditionIndex2)

		'Set FormatCondition Object formula
		fc2.Formula1 = "=IF(SUM(B1:B2)>100,TRUE,FALSE)"

		'Set FormatCondition Object Background Color
		fc2.Style.BackgroundColor = Color.Red

		sheet.Cells("B3").Formula = "=SUM(B1:B2)"

		'Put Value in Cell C4 
		sheet.Cells("C4").PutValue("If Sum of B1:B2 is greater than 100, B3 will have RED background")

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "ConditionalFormatting.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "ConditionalFormatting.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub

End Class



