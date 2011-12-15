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
Imports Aspose.Cells.Tables

Namespace Aspose.Cells.Demos

	Partial Public Class Format_List_Object
		Inherits System.Web.UI.Page
		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
			CreateStaticReport()
		End Sub

		Public Shared Sub CreateStaticReport()
			'Create a workbook.
			Dim workbook As New Workbook()

			'Obtaining the reference of the newly added worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Obtaining Worksheet's cells collection
			Dim cells As Cells = sheet.Cells

			'Initialize cell and Set the value to the cells
			'Set Value for Column Header (first row)
			Dim cell As Aspose.Cells.Cell = cells("A1")
			cell.PutValue("Employee")
			cell = cells("B1")
			cell.PutValue("Quarter")
			cell = cells("C1")
			cell.PutValue("Product")
			cell = cells("D1")
			cell.PutValue("Continent")
			cell = cells("E1")
			cell.PutValue("Country")
			cell = cells("F1")
			cell.PutValue("Sale")

			'Set Value for row cells in Column A
			cell = cells("A2")
			cell.PutValue("David")
			cell = cells("A3")
			cell.PutValue("David")
			cell = cells("A4")
			cell.PutValue("David")
			cell = cells("A5")
			cell.PutValue("David")
			cell = cells("A6")
			cell.PutValue("James")
			cell = cells("A7")
			cell.PutValue("James")
			cell = cells("A8")
			cell.PutValue("James")
			cell = cells("A9")
			cell.PutValue("James")
			cell = cells("A10")
			cell.PutValue("James")
			cell = cells("A11")
			cell.PutValue("Miya")
			cell = cells("A12")
			cell.PutValue("Miya")
			cell = cells("A13")
			cell.PutValue("Miya")
			cell = cells("A14")
			cell.PutValue("Miya")
			cell = cells("A15")
			cell.PutValue("Miya")

			'Set Value for row cells in Column B
			cell = cells("B2")
			cell.PutValue(1)
			cell = cells("B3")
			cell.PutValue(2)
			cell = cells("B4")
			cell.PutValue(3)
			cell = cells("B5")
			cell.PutValue(4)
			cell = cells("B6")
			cell.PutValue(1)
			cell = cells("B7")
			cell.PutValue(2)
			cell = cells("B8")
			cell.PutValue(3)
			cell = cells("B9")
			cell.PutValue(4)
			cell = cells("B10")
			cell.PutValue(4)
			cell = cells("B11")
			cell.PutValue(1)
			cell = cells("B12")
			cell.PutValue(1)
			cell = cells("B13")
			cell.PutValue(2)
			cell = cells("B14")
			cell.PutValue(2)
			cell = cells("B15")
			cell.PutValue(2)

			'Set Value for row cells in Column C
			cell = cells("C2")
			cell.PutValue("Maxilaku")
			cell = cells("C3")
			cell.PutValue("Maxilaku")
			cell = cells("C4")
			cell.PutValue("Chai")
			cell = cells("C5")
			cell.PutValue("Maxilaku")
			cell = cells("C6")
			cell.PutValue("Chang")
			cell = cells("C7")
			cell.PutValue("Chang")
			cell = cells("C8")
			cell.PutValue("Chang")
			cell = cells("C9")
			cell.PutValue("Chang")
			cell = cells("C10")
			cell.PutValue("Chang")
			cell = cells("C11")
			cell.PutValue("Geitost")
			cell = cells("C12")
			cell.PutValue("Chai")
			cell = cells("C13")
			cell.PutValue("Geitost")
			cell = cells("C14")
			cell.PutValue("Geitost")
			cell = cells("C15")
			cell.PutValue("Geitost")

			'Set Value for row cells in Column D
			cell = cells("D2")
			cell.PutValue("Asia")
			cell = cells("D3")
			cell.PutValue("Asia")
			cell = cells("D4")
			cell.PutValue("Asia")
			cell = cells("D5")
			cell.PutValue("Asia")
			cell = cells("D6")
			cell.PutValue("Europe")
			cell = cells("D7")
			cell.PutValue("Europe")
			cell = cells("D8")
			cell.PutValue("Europe")
			cell = cells("D9")
			cell.PutValue("Europe")
			cell = cells("D10")
			cell.PutValue("Europe")
			cell = cells("D11")
			cell.PutValue("America")
			cell = cells("D12")
			cell.PutValue("America")
			cell = cells("D13")
			cell.PutValue("America")
			cell = cells("D14")
			cell.PutValue("America")
			cell = cells("D15")
			cell.PutValue("America")

			'Set Value for row cells in Column E
			cell = cells("E2")
			cell.PutValue("China")
			cell = cells("E3")
			cell.PutValue("India")
			cell = cells("E4")
			cell.PutValue("Korea")
			cell = cells("E5")
			cell.PutValue("India")
			cell = cells("E6")
			cell.PutValue("France")
			cell = cells("E7")
			cell.PutValue("France")
			cell = cells("E8")
			cell.PutValue("Germany")
			cell = cells("E9")
			cell.PutValue("Italy")
			cell = cells("E10")
			cell.PutValue("France")
			cell = cells("E11")
			cell.PutValue("U.S.")
			cell = cells("E12")
			cell.PutValue("U.S.")
			cell = cells("E13")
			cell.PutValue("Brazil")
			cell = cells("E14")
			cell.PutValue("U.S.")
			cell = cells("E15")
			cell.PutValue("U.S.")

			'Set Value for row cells in Column F
			cell = cells("F2")
			cell.PutValue(2000)
			cell = cells("F3")
			cell.PutValue(500)
			cell = cells("F4")
			cell.PutValue(1200)
			cell = cells("F5")
			cell.PutValue(1500)
			cell = cells("F6")
			cell.PutValue(500)
			cell = cells("F7")
			cell.PutValue(1500)
			cell = cells("F8")
			cell.PutValue(800)
			cell = cells("F9")
			cell.PutValue(900)
			cell = cells("F10")
			cell.PutValue(500)
			cell = cells("F11")
			cell.PutValue(1600)
			cell = cells("F12")
			cell.PutValue(600)
			cell = cells("F13")
			cell.PutValue(2000)
			cell = cells("F14")
			cell.PutValue(500)
			cell = cells("F15")
			cell.PutValue(900)

			'Adding a new LIst Object to the worksheet
			Dim listObject As ListObject = sheet.ListObjects(sheet.ListObjects.Add("A1", "F15", True))

			'Adding Default Style to the Worksheet
			listObject.TableStyleType = TableStyleType.TableStyleMedium9

			'Show Total
			listObject.ShowTotals = True

			'Saving the Excel file 
			workbook.Save(HttpContext.Current.Response, "List_Object.xlsx", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Xlsx))

			'End response to avoid unneeded html after xls
			HttpContext.Current.Response.End()
		End Sub
	End Class
End Namespace