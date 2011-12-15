Imports Microsoft.VisualBasic
Imports System

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for SalesTotals.
	''' </summary>
	Public Class SalesTotals
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)
		End Sub

		Public Function CreateSalesTotals() As Workbook
			Try
				DBInit()
			Catch
			Finally
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try

			'Open template file
			Dim designerFile As String = MapPath("~/Designer/Northwind.xls")
			Dim workbook As New Workbook(designerFile)

			Try
				'Specify SQL and execute the query to fill the datatable
				Me.oleDbSelectCommand1.CommandText = "SELECT [Order Subtotals].Subtotal, [Order Subtotals].OrderID, " & ControlChars.CrLf & "				Customers.CompanyName, Customers.CustomerID FROM Customers " & ControlChars.CrLf & "				INNER JOIN ([Order Subtotals] INNER JOIN Orders ON [Order Subtotals].OrderID = Orders.OrderID) " & ControlChars.CrLf & "				ON Customers.CustomerID = Orders.CustomerID " & ControlChars.CrLf & "				WHERE (Orders.ShippedDate BETWEEN #1/1/1995# AND #12/31/1995#) AND ([Order Subtotals].Subtotal > 2500) " & ControlChars.CrLf & "				ORDER BY [Order Subtotals].Subtotal DESC"
				Me.oleDbDataAdapter1.Fill(Me.dataTable1)
			Catch
			Finally
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try

			'Get the worksheet
			Dim sheet As Worksheet = workbook.Worksheets("Sheet12")
			'Name the worksheet
			sheet.Name = "Sales Totals"
			'Get the cells
			Dim cells As Cells = sheet.Cells
			'Import the datatable to the sheet
			cells.ImportDataTable(Me.dataTable1, False, 3, 1, Me.dataTable1.Rows.Count, 3)

			Dim totalSum As Decimal = 0.0D
			'Input some value to the cells
			For i As Integer = 0 To Me.dataTable1.Rows.Count - 1
				totalSum += CDec(Me.dataTable1.Rows(i)("Subtotal"))
				cells(3 + i, 5).PutValue(i + 1)
			Next i

			'Input value and create a style to apply it
			cells(3 + Me.dataTable1.Rows.Count, 0).PutValue("Total:")
			Dim style As Style = workbook.Styles(workbook.Styles.Add())
			style.Font.IsBold = True
			cells(3 + Me.dataTable1.Rows.Count, 0).SetStyle(style)
			'Input a value
			cells(3 + Me.dataTable1.Rows.Count, 1).PutValue(CDbl(totalSum))
			'Remove the unnecessary worksheets in the workbook
            Dim iworkbook As Integer = 0
            Do While iworkbook < workbook.Worksheets.Count
                sheet = workbook.Worksheets(iworkbook)
                If sheet.Name <> "Sales Totals" Then
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
