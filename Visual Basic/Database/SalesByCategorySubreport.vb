Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Configuration
Imports System.Collections
Imports System.Web
Imports System.Web.Security
'using System;

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for SalesByCategorySubreport.
	''' </summary>
	Public Class SalesByCategorySubreport
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)

		End Sub

		Public Function CreateSalesByCategorySubreport() As Workbook
			Try
				DBInit()
			Catch
			Finally
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try

			'Open the template file
		Dim designerFile As String = MapPath("~/Designer/Northwind.xls")
			Dim workbook As New Workbook(designerFile)

			Try
				'Specify SQL and execute the query to fill a datatable
				Me.oleDbSelectCommand1.CommandText = "SELECT DISTINCTROW Products.ProductName, Sum([Order Details Extended].ExtendedPrice) AS ProductSales" & ControlChars.CrLf & "				FROM Categories INNER JOIN (Products INNER JOIN (Orders INNER JOIN [Order Details Extended] ON Orders.OrderID = [Order Details Extended].OrderID) ON Products.ProductID = [Order Details Extended].ProductID) ON Categories.CategoryID = Products.CategoryID" & ControlChars.CrLf & "				WHERE (((Orders.OrderDate) Between #1/1/1995# And #12/31/1995#))" & ControlChars.CrLf & "				GROUP BY Categories.CategoryID, Categories.CategoryName, Products.ProductName" & ControlChars.CrLf & "				ORDER BY Products.ProductName"
				Me.oleDbDataAdapter1.Fill(Me.dataTable1)
			Catch
			Finally
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try

			'Get a worksheet
			Dim sheet As Worksheet = workbook.Worksheets("Sheet9")
			'Name the sheet
			sheet.Name = "Sales By Category Subreport"
			'Get the cells collection
			Dim cells As Cells = sheet.Cells
			'Import the datatable to the sheet
			cells.ImportDataTable(Me.dataTable1, False, 0, 0)
			'Remove the unnecessary worksheets
			Dim i As Integer = 0
			Do While i < workbook.Worksheets.Count
				sheet = workbook.Worksheets(i)
				If sheet.Name <> "Sales By Category Subreport" Then
					workbook.Worksheets.RemoveAt(i)
					i -= 1
				End If
				i += 1
			Loop
			'Get the generated workbook
			Return workbook
		End Function

	End Class
End Namespace


