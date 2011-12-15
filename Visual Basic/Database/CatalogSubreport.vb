Imports Microsoft.VisualBasic
Imports System

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for CatalogSubreport.
	''' </summary>
	Public Class CatalogSubreport
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)

		End Sub

		Public Function CreateCatalogSubreport() As Workbook
			Try
				DBInit()

				'Open the connection object
				Me.oleDbConnection1.Open()
				'Specify an SQL as command text
				Me.oleDbSelectCommand1.CommandText = "SELECT ProductName, ProductID, QuantityPerUnit, UnitPrice FROM Products ORDER BY " & "ProductName"
				'Fill a datatable
				Me.oleDbDataAdapter1.Fill(Me.dataTable1)
			Catch
			Finally
				If Me.oleDbDataAdapter1 IsNot Nothing Then
					Me.oleDbDataAdapter1.Dispose()
				End If
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try

			'Open a template file
		Dim designerFile As String = MapPath("~/Designer/Northwind.xls")
		Dim workbook As New Workbook(designerFile)

			'Get the sheet
			Dim sheet As Worksheet = workbook.Worksheets("Sheet3")
			'Name the sheet
			sheet.Name = "Catalog Subreport"
			'Get the cells in the sheet
			Dim cells As Cells = sheet.Cells
			'Import the datatable to the sheet
			cells.ImportDataTable(Me.dataTable1, False, 0, 1)

			'Remove the unnecessary worksheets in the workbook
			Dim i As Integer = 0
			Do While i < workbook.Worksheets.Count
				sheet = workbook.Worksheets(i)
				If sheet.Name <> "Catalog Subreport" Then
					workbook.Worksheets.RemoveAt(i)
					i -= 1
				End If

				i += 1
			Loop
			'Retrun the generated workbook
			Return workbook
		End Function
	End Class
End Namespace


