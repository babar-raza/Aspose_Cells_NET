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
Imports Aspose.Cells.Pivot
Imports System.Drawing

Namespace Aspose.Cells.Demos

	Partial Public Class Pivot_Table
		Inherits System.Web.UI.Page
		Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
			CreateStaticReport()
		End Sub

		Public Sub CreateStaticReport()
			'Instantiating an Workbook object
			Dim workbook As New Workbook()

			'Obtaining the reference of the newly added worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)
			Dim cells As Cells = sheet.Cells

			'Setting the value to the cells
			Dim cell As Cell = cells("A1")
			cell.PutValue("Sport")
			cell = cells("B1")
			cell.PutValue("Quarter")
			cell = cells("C1")
			cell.PutValue("Sales")


			cell = cells("A2")
			cell.PutValue("Golf")
			cell = cells("A3")
			cell.PutValue("Golf")
			cell = cells("A4")
			cell.PutValue("Tennis")
			cell = cells("A5")
			cell.PutValue("Tennis")
			cell = cells("A6")
			cell.PutValue("Tennis")
			cell = cells("A7")
			cell.PutValue("Tennis")
			cell = cells("A8")
			cell.PutValue("Golf")


			cell = cells("B2")
			cell.PutValue("Qtr3")
			cell = cells("B3")
			cell.PutValue("Qtr4")
			cell = cells("B4")
			cell.PutValue("Qtr3")
			cell = cells("B5")
			cell.PutValue("Qtr4")
			cell = cells("B6")
			cell.PutValue("Qtr3")
			cell = cells("B7")
			cell.PutValue("Qtr4")
			cell = cells("B8")
			cell.PutValue("Qtr3")

			cell = cells("C2")
			cell.PutValue(1500)
			cell = cells("C3")
			cell.PutValue(2000)
			cell = cells("C4")
			cell.PutValue(600)
			cell = cells("C5")
			cell.PutValue(1500)
			cell = cells("C6")
			cell.PutValue(4070)
			cell = cells("C7")
			cell.PutValue(5000)
			cell = cells("C8")
			cell.PutValue(6430)

			Dim pivotTables As PivotTableCollection = sheet.PivotTables

			'Adding a PivotTable to the worksheet
			Dim index As Integer = pivotTables.Add("=A1:C8", "E20", "PivotTable1")

			'Accessing the instance of the newly added PivotTable
			Dim pivotTable As PivotTable = pivotTables(index)

			'Draging the first field to the row area.
			pivotTable.AddFieldToArea(PivotFieldType.Row, 0)

			'Draging the second field to the column area.
			pivotTable.AddFieldToArea(PivotFieldType.Column, 1)

			'Draging the third field to the data area.
			pivotTable.AddFieldToArea(PivotFieldType.Data, 2)

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "PivotTable.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "PivotTable.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub
	End Class
End Namespace