Imports Microsoft.VisualBasic
Imports System

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for SalesByYear.
	''' </summary>
	Public Class SalesByYear
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)

		End Sub

		Public Function CreateSalesByYear() As Workbook
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

			'Specify an SQL and execute the query to fill a datatable
			Me.oleDbSelectCommand1.CommandText = "SELECT DISTINCTROW Format([ShippedDate],""yyyy-mm-dd"") AS [ShippedDate], Orders.OrderID, [Order Subtotals].Subtotal as Subtotal" & ControlChars.CrLf & "				FROM Orders INNER JOIN [Order Subtotals] ON Orders.OrderID = [Order Subtotals].OrderID" & ControlChars.CrLf & "				WHERE( (Orders.ShippedDate) Is Not Null)"
			Me.oleDbDataAdapter1.Fill(Me.dataTable1)

			'Get the sheet
			Dim sheet As Worksheet = workbook.Worksheets("Sheet10")
			'Name the sheet
			sheet.Name = "Sales By Year"
			'Get the cells collection
			Dim cells As Cells = sheet.Cells
			'Import the datatable to the sheet
			cells.ImportDataTable(Me.dataTable1, False, 6, 2)
			'Input values to some cells
			For i As Integer = 0 To Me.dataTable1.Rows.Count - 1
				cells(6 + i, 1).PutValue(i + 1)
			Next i
			'Remove the unnecessary worksheets in the workbook
            Dim iworkbook As Integer = 0
            Do While iworkbook < workbook.Worksheets.Count
                sheet = workbook.Worksheets(iworkbook)
                If sheet.Name <> "Sales By Year" Then
                    workbook.Worksheets.RemoveAt(iworkbook)
                    iworkbook -= 1
                End If
                iworkbook += 1
            Loop
			'Get the generated workbook
			Return workbook
		End Function

	End Class
End Namespace


