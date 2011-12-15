Imports Microsoft.VisualBasic
Imports System

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for ProductsList.
	''' </summary>
	Public Class ProductsList
		Inherits DbBase


		Public Sub New(ByVal path As String)
			MyBase.New(path)
		End Sub


		Public Function CreateProductsList() As Workbook
			Try
				DBInit()
				'Open the connection
				Me.oleDbConnection1.Open()
				'Specify an SQL query as command text
				Me.oleDbSelectCommand1.CommandText = "SELECT	DISTINCTROW Products.ProductName, " & ControlChars.CrLf & "														Categories.CategoryName, " & ControlChars.CrLf & "														Products.QuantityPerUnit, " & ControlChars.CrLf & "														Products.UnitsInStock" & ControlChars.CrLf & "													FROM Categories INNER JOIN Products " & ControlChars.CrLf & "													ON	Categories.CategoryID = Products.CategoryID" & ControlChars.CrLf & "													WHERE" & ControlChars.CrLf & "														(((Products.Discontinued) = No))" & ControlChars.CrLf & "													Order by Products.ProductName"
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

			'Open a template excel file
		Dim designerFile As String = MapPath("~/Designer/Northwind.xls")
			Dim workbook As New Workbook(designerFile)

			'Get the first worksheet in the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Import a datatable to the sheet
			sheet.Cells.ImportDataTable(Me.dataTable1, False, 6, 1)
			'Name the sheet
			sheet.Name = "Products List"

			'Remove all other worksheets (except the first worksheet) in the workbook
			Do While workbook.Worksheets.Count > 1

				workbook.Worksheets.RemoveAt(workbook.Worksheets.Count - 1)
			Loop

			'Return the generated workbook
			Return workbook
		End Function


	End Class
End Namespace


