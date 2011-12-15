Imports Microsoft.VisualBasic
Imports System

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for CustomerLabels.
	''' </summary>
	Public Class CustomerLabels
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)

		End Sub

		Public Function CreateCustomerLabels() As Workbook
			Try
				DBInit()
				'Open the connection object
				Me.oleDbConnection1.Open()
				'Specify SQL as command text
				Me.oleDbSelectCommand1.CommandText = "SELECT CompanyName, Address, City, Region, PostalCode, Country, CustomerID FROM " & "Customers ORDER BY Country, Region"
				'Fill the datatable
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

			'Get a worksheet
			Dim sheet As Worksheet = workbook.Worksheets("Sheet4")
			'Name the worksheet
			sheet.Name = "Customer Labels"
			'Get the cells collection in the worksheet
			Dim cells As Cells = sheet.Cells
			Dim row As Integer = 0
			Dim column As Byte = 0
			For i As Integer = 0 To Me.dataTable1.Rows.Count - 1
				Dim remainder As Integer = i Mod 3
				Dim cell As Cell
				Select Case remainder
					Case 0
						column = 0
					Case 1
						column = 3
					Case 2
						column = 6
				End Select
				'Get a cell
				cell = cells(row, column)
				'Put a value into it
				cell.PutValue(CStr(Me.dataTable1.Rows(i)("CompanyName")))
				'Get another cell
				cell = cells(row + 1, column)
				'Put a value into it
				cell.PutValue(CStr(Me.dataTable1.Rows(i)("Address")))
				'Get another cell
				cell = cells(row + 2, column)
				Dim contact As String = ""

				If Me.dataTable1.Rows(i)("City") IsNot DBNull.Value Then
					contact &= CStr(Me.dataTable1.Rows(i)("City")) & " "
				End If
				If Me.dataTable1.Rows(i)("Region") IsNot DBNull.Value Then
					contact &= CStr(Me.dataTable1.Rows(i)("Region")) & " "
				End If
				If Me.dataTable1.Rows(i)("PostalCode") IsNot DBNull.Value Then
					contact &= CStr(Me.dataTable1.Rows(i)("PostalCode"))
				End If

				'Put the value to it
				cell.PutValue(contact)
				'Get another cell
				cell = cells(row + 3, column)
				'Put a value to it
				cell.PutValue(CStr(Me.dataTable1.Rows(i)("Country")))

				If remainder = 2 Then
					row += 5
				End If

			Next i

			'Remove unnecessary worksheets in the workbook
            Dim iworkbook As Integer = 0
            Do While iworkbook < workbook.Worksheets.Count
                sheet = workbook.Worksheets(iworkbook)
                If sheet.Name <> "Customer Labels" Then
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


