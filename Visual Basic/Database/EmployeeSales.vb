Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for EmployeeSales.
	''' </summary>
	Public Class EmployeeSales
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)

		End Sub

		Public Function CreateEmployeeSales() As Workbook
			Try
				DBInit()
			Catch
			Finally
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try


			'Open a template file
		Dim designerFile As String = MapPath("~/Designer/Northwind.xls")
			Dim workbook As New Workbook(designerFile)

			'Get the worksheet
			Dim sheet As Worksheet = workbook.Worksheets("Sheet5")
			'Name the worksheet
			sheet.Name = "Employee Sales UK"
			'Get the cells
			Dim cells As Cells = sheet.Cells
			'Get the worksheet
			sheet = workbook.Worksheets("Sheet6")
			'Get its cells
			cells = sheet.Cells
			'Name the sheet
			sheet.Name = "Employee Sales USA"

			ReadEmployees()
			'Create datatable array
			Dim dtSales() As DataTable = Me.CreateDataResult()

			Dim currentUKRow As Integer = 6
			Dim currentUSARow As Integer = 6

			Dim styleIndex As Integer
			Dim style As Style
			'Create a header style with specific formatting
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Double
			style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Double
			style.Borders.SetColor(Color.Black)
			style.Font.Size = 12
			style.Font.IsBold = True
			style.IsTextWrapped = True
			style.HorizontalAlignment = TextAlignmentType.Center
			style.Name = "HeaderStyle"

			'Create different styles with specific formattings and apply to different cells
			'Input different values and set formulas to some cells, importing datatables to the cells  
			For i As Integer = 0 To Me.dataTable1.Rows.Count - 1
				Dim employeeName As String = CStr(Me.dataTable1.Rows(i)("LastName")) & "," & CStr(Me.dataTable1.Rows(i)("FirstName"))
				If Me.dataTable1.Rows(i)("Country").ToString() = "UK" Then
					sheet = workbook.Worksheets("Employee Sales UK")
					cells = sheet.Cells

					cells(currentUKRow - 2, 0).PutValue("Salesperson:" & employeeName)
					style = workbook.Styles(workbook.Styles.Add())
					style.Font.IsBold = True
					style.Font.Size = 12

					cells(currentUKRow - 2, 0).SetStyle(style)

					If CDec(Me.dataTable1.Rows(i)("TotalSales")) > 5000 Then
						cells(currentUKRow - 2, 3).PutValue("Exceeded Goal!")
						style = workbook.Styles(workbook.Styles.Add())

						Dim font As Font = style.Font
						font.Color = Color.Red
						font.IsItalic = True
						font.Size = 12
						font.IsBold = True

						cells(currentUKRow - 2, 3).SetStyle(style)
					End If
					cells.SetRowHeight(currentUKRow - 2, 19)
					cells.SetRowHeight(currentUKRow - 1, 4)
					cells.SetRowHeight(currentUKRow, 48)

					style = workbook.Styles("HeaderStyle")
					For j As Integer = 1 To 4
						cells(currentUKRow, CByte(j)).SetStyle(style)
					Next j
					cells(currentUKRow, 1).PutValue("Order ID:")
					cells(currentUKRow, 2).PutValue("Sales Amount:")
					cells(currentUKRow, 3).PutValue("Percent of Salesperson's Total:")
					cells(currentUKRow, 4).PutValue("Percent of Country Total:")
					currentUKRow += 1

					cells.ImportDataTable(dtSales(i), False, currentUKRow, 1)
					Dim startCellName1 As String = CellsHelper.CellIndexToName(currentUKRow, 2)
					Dim startCellName2 As String = CellsHelper.CellIndexToName(currentUKRow, 4)

					currentUKRow += dtSales(i).Rows.Count - 1
					Dim endCellName1 As String = CellsHelper.CellIndexToName(currentUKRow, 2)
					Dim endCellName2 As String = CellsHelper.CellIndexToName(currentUKRow, 4)

					cells(currentUKRow + 1, 2).Formula = "=sum(" & startCellName1 & ":" & endCellName1 & ")"

					cells(currentUKRow + 1, 4).Formula = "=sum(" & startCellName2 & ":" & endCellName2 & ")"

					currentUKRow += 4

				Else
					sheet = workbook.Worksheets("Employee Sales USA")
					cells = sheet.Cells

					cells(currentUSARow - 2, 0).PutValue("Salesperson:" & employeeName)

					style = workbook.Styles(workbook.Styles.Add())
					style.Font.IsBold = True
					style.Font.Size = 12
					cells(currentUSARow - 2, 0).SetStyle(style)

					If CDec(Me.dataTable1.Rows(i)("TotalSales")) > 5000 Then
						style = workbook.Styles(workbook.Styles.Add())

						cells(currentUSARow - 2, 3).PutValue("Exceeded Goal!")
						Dim font As Font = style.Font
						font.Color = Color.Red
						font.IsItalic = True
						font.Size = 12
						font.IsBold = True

						cells(currentUSARow - 2, 3).SetStyle(style)
					End If
					cells.SetRowHeight(currentUSARow - 2, 19)
					cells.SetRowHeight(currentUSARow - 1, 4)
					cells.SetRowHeight(currentUSARow, 48)

					style = workbook.Styles("HeaderStyle")
					For j As Integer = 1 To 4
						cells(currentUSARow, CByte(j)).SetStyle(style)
					Next j
					cells(currentUSARow, 1).PutValue("Order ID:")
					cells(currentUSARow, 2).PutValue("Sales Amount:")
					cells(currentUSARow, 3).PutValue("Percent of Salesperson's Total:")
					cells(currentUSARow, 4).PutValue("Percent of Country Total:")
					currentUSARow += 1

					cells.ImportDataTable(dtSales(i), False, currentUSARow, 1)

					Dim startCellName1 As String = CellsHelper.CellIndexToName(currentUSARow, 2)
					Dim startCellName2 As String = CellsHelper.CellIndexToName(currentUSARow, 4)

					currentUSARow += dtSales(i).Rows.Count - 1
					Dim endCellName1 As String = CellsHelper.CellIndexToName(currentUSARow, 2)
					Dim endCellName2 As String = CellsHelper.CellIndexToName(currentUSARow, 4)

					cells(currentUSARow + 1, 2).Formula = "=sum(" & startCellName1 & ":" & endCellName1 & ")"

					cells(currentUSARow + 1, 4).Formula = "=sum(" & startCellName2 & ":" & endCellName2 & ")"

					currentUSARow += 4
				End If
			Next i
			'Remove unnecessary worksheets
            Dim iworkbook As Integer = 0
            Do While iworkbook < workbook.Worksheets.Count
                sheet = workbook.Worksheets(iworkbook)
                If sheet.Name <> "Employee Sales UK" AndAlso sheet.Name <> "Employee Sales USA" Then
                    workbook.Worksheets.RemoveAt(iworkbook)
                    iworkbook -= 1
                End If

                iworkbook += 1
            Loop
			'Get the generated workbook
			Return workbook
		End Function

		Private Sub ReadEmployees()
			Try
				'Open the connection
				Me.oleDbConnection1.Open()
				'Specify SQL
				Me.oleDbSelectCommand1.CommandText = "SELECT Country, EmployeeID, FirstName, LastName FROM Employees ORDER BY Country, " & "LastName, FirstName"
				'Fill a datatable executing the query
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
			'Add a column to the datatable
			Me.dataTable1.Columns.Add("TotalSales", GetType(Decimal))

		End Sub

		Private Function CreateDataResult() As DataTable()
			'Create datatable array
			Dim dtSales(Me.dataTable1.Rows.Count - 1) As DataTable
			Dim totalUKSales As Decimal = 0.0D
			Dim totalUSASales As Decimal = 0.0D
			'Specify SQL
			Dim cmd As String = "SELECT Orders.OrderID, [Order Subtotals].Subtotal as SaleAmount FROM Employees INNER " & "JOIN (Orders INNER JOIN [Order Subtotals] ON Orders.OrderID = [Order Subtotals].OrderID) " & "ON Employees.EmployeeID = Orders.EmployeeID"

			'Create different datatables and fill them with different sets of data based on specific SQL
			For i As Integer = 0 To Me.dataTable1.Rows.Count - 1
				dtSales(i) = New DataTable()
				dtSales(i).Columns.Add("OrderID", GetType(Integer))
				dtSales(i).Columns.Add("SaleAmount", GetType(Decimal))
				dtSales(i).Columns.Add("PercentOfPerson", GetType(Decimal))
				dtSales(i).Columns.Add("PercentOfCountry", GetType(Decimal))

				Dim totalPersonSales As Decimal = 0.0D
				Try
					Me.oleDbDataAdapter2 = New OleDbDataAdapter()
					Dim cmdText As String = cmd & " where Employees.EmployeeID =" & Me.dataTable1.Rows(i)("EmployeeID").ToString()
					Me.oleDbDataAdapter2.SelectCommand = New OleDbCommand(cmdText, Me.oleDbConnection1)
					Me.oleDbConnection1.Open()
					Me.oleDbDataAdapter2.Fill(dtSales(i))
				Catch
				Finally
					If Me.oleDbDataAdapter2 IsNot Nothing Then
						Me.oleDbDataAdapter2.Dispose()
					End If
					If Me.oleDbConnection1 IsNot Nothing Then
						Me.oleDbConnection1.Close()
					End If
				End Try

				'Get total sales amount of a salesperson
				For j As Integer = 0 To dtSales(i).Rows.Count - 1
					totalPersonSales += CDec(dtSales(i).Rows(j)("SaleAmount"))
				Next j

				Me.dataTable1.Rows(i)("TotalSales") = totalPersonSales

				'Get the percent
				For j As Integer = 0 To dtSales(i).Rows.Count - 1
					dtSales(i).Rows(j)("PercentOfPerson") = CDec(dtSales(i).Rows(j)("SaleAmount")) / totalPersonSales
				Next j
				If Me.dataTable1.Rows(i)("Country").ToString() = "UK" Then
					totalUKSales += totalPersonSales
				Else
					totalUSASales += totalPersonSales
				End If
			Next i
			For i As Integer = 0 To dtSales.Length - 1
				For j As Integer = 0 To dtSales(i).Rows.Count - 1
					If Me.dataTable1.Rows(i)("Country").ToString() = "UK" Then
						dtSales(i).Rows(j)("PercentOfCountry") = CDec(dtSales(i).Rows(j)("SaleAmount")) / totalUKSales
					Else
						dtSales(i).Rows(j)("PercentOfCountry") = CDec(dtSales(i).Rows(j)("SaleAmount")) / totalUSASales
					End If
				Next j
			Next i
			'Get the generated datatables
			Return dtSales
		End Function

	End Class
End Namespace


