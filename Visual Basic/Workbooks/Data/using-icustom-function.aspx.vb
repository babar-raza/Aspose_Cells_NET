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

Partial Public Class Workbooks_Data_UsingICustomFunction
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		'Call Method to create report
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Open the workbook
		Dim workbook As New Workbook()

		'Obtaining the reference of the first worksheet
		Dim worksheet As Worksheet = workbook.Worksheets(0)

		'Adding a sample value to "A1" cell
		worksheet.Cells("B1").PutValue(5)

		'Adding a sample value to "A2" cell
		worksheet.Cells("C1").PutValue(100)

		'Adding a sample value to "A3" cell
		worksheet.Cells("C2").PutValue(150)

		'Adding a sample value to "B1" cell
		worksheet.Cells("C3").PutValue(60)

		'Adding a sample value to "B2" cell
		worksheet.Cells("C4").PutValue(32)

		'Adding a sample value to "B2" cell
		worksheet.Cells("C5").PutValue(62)

		'Adding custom formula to Cell A1
		workbook.Worksheets(0).Cells("A1").Formula = "=MyFunc(B1,C1:C5)"

		'Calcualting Formulas
		workbook.CalculateFormula(False, New CustomFunction())

		'Assign resultant value to Cell A1
		workbook.Worksheets(0).Cells("A1").PutValue(workbook.Worksheets(0).Cells("A1").Value)

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "UsingICustomFunction.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "UsingICustomFunction.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub
End Class

Public Class CustomFunction
	Implements ICustomFunction

	Public Function CalculateCustomFunction(ByVal functionName As String, ByVal paramsList As System.Collections.ArrayList, ByVal contextObjects As System.Collections.ArrayList) As Object Implements ICustomFunction.CalculateCustomFunction
		'get value of first parameter
		Dim firstParamB1 As Decimal = System.Convert.ToDecimal(paramsList(0))

		'get value of second parameter
		Dim secondParamC1C5 As Array = CType(paramsList(1), Array)

		Dim total As Decimal = 0D

		' get every item value of second parameter
		For Each value As Object() In secondParamC1C5
			total += System.Convert.ToDecimal(value(0))
		Next value

		total = total / firstParamB1

		'return result of the function
		Return total
	End Function
End Class

