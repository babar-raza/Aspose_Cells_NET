Imports Microsoft.VisualBasic
Imports System

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for SalesByYearSubreport.
	''' </summary>
	Public Class SalesByYearSubreport
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)
		End Sub

		Public Function CreateSalesByYearSubreport() As Workbook
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
				'Specify an SQL and execute the query to fill the datatable
				Me.oleDbSelectCommand1.CommandText = "SELECT" & ControlChars.CrLf & "					DISTINCTROW COUNT(Orders.OrderID) AS Orders, " & ControlChars.CrLf & "					SUM([Order Subtotals].Subtotal) AS Sales, " & ControlChars.CrLf & "					FORMAT(ORDERS.SHIPPEDDATE, " & ControlChars.CrLf & "					'yyyy/Q') AS Quarter" & ControlChars.CrLf & "				FROM" & ControlChars.CrLf & "					Orders INNER JOIN [Order Subtotals] " & ControlChars.CrLf & "				ON" & ControlChars.CrLf & "					Orders.OrderID = [Order Subtotals].OrderID" & ControlChars.CrLf & "				WHERE" & ControlChars.CrLf & "					(orders.shippeddate IS NOT NULL) GROUP BY FORMAT(ORDERS.SHIPPEDDATE ,  'yyyy/Q')"
				Me.oleDbDataAdapter1.Fill(Me.dataTable1)
			Catch
			Finally
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try

			'Get the worksheet
			Dim sheet As Worksheet = workbook.Worksheets("Sheet11")
			'Set its name
			sheet.Name = "Sales By Year Subreport"
			'Get the cells
			Dim cells As Cells = sheet.Cells

			Dim currentRow As Integer = 0
			Dim totalOrders As Integer = 0
			Dim totalSales As Decimal = 0.0D
			Dim thisYear As String = ""
			SetSalesByYearSubreportStyles(workbook)
			For i As Integer = 0 To Me.dataTable1.Rows.Count - 1
				If i = 0 Then
					thisYear = Me.dataTable1.Rows(0)("Quarter").ToString().Substring(0, 4)
					CreateSalesByYearSubreportHeader(workbook, cells, 0, thisYear)
					CreateData(cells, 2, 0)
					totalOrders += CInt(Fix(Me.dataTable1.Rows(0)("Orders")))
					totalSales += CDec(Me.dataTable1.Rows(0)("Sales"))
					currentRow = 3
				Else
					If thisYear = Me.dataTable1.Rows(i)("Quarter").ToString().Substring(0, 4) Then
						CreateData(cells, currentRow, i)
						totalOrders += CInt(Fix(Me.dataTable1.Rows(i)("Orders")))
						totalSales += CDec(Me.dataTable1.Rows(i)("Sales"))
						currentRow += 1
						If i = Me.dataTable1.Rows.Count - 1 Then
							CreateFooter(workbook, cells, currentRow, totalOrders, totalSales)
						End If
					Else
						CreateFooter(workbook, cells, currentRow, totalOrders, totalSales)
						totalOrders = 0
						totalSales = 0.0D
						currentRow += 1
						thisYear = Me.dataTable1.Rows(i)("Quarter").ToString().Substring(0, 4)
						If i <> Me.dataTable1.Rows.Count - 1 Then
							CreateSalesByYearSubreportHeader(workbook, cells, currentRow, thisYear)
							currentRow += 2
							CreateData(cells, currentRow, i)
							totalOrders += CInt(Fix(Me.dataTable1.Rows(i)("Orders")))
							totalSales += CDec(Me.dataTable1.Rows(i)("Sales"))
							currentRow += 1
						End If
					End If
				End If
			Next i
			'Remove the unnecessary worksheets in the workbook
            Dim iworkbook As Integer = 0
            Do While iworkbook < workbook.Worksheets.Count
                sheet = workbook.Worksheets(iworkbook)
                If sheet.Name <> "Sales By Year Subreport" Then
                    workbook.Worksheets.RemoveAt(iworkbook)
                    iworkbook -= 1
                End If
                iworkbook += 1
            Loop
			'Get the generated workbook
			Return workbook
		End Function

		Private Sub CreateFooter(ByVal workbook As Workbook, ByVal cells As Cells, ByVal startRow As Integer, ByVal totalOrders As Integer, ByVal totalSales As Decimal)

			'Get the style
			Dim style As Style = workbook.Styles("Bold")
			'Put value and apply style
			cells(startRow, 1).PutValue("Totals:")
			cells(startRow, 1).SetStyle(style)
			'Put values to cells
			cells(startRow, 2).PutValue(totalOrders)
			cells(startRow, 3).PutValue(CDbl(totalSales))
		End Sub

		Private Sub CreateData(ByVal cells As Cells, ByVal startRow As Integer, ByVal index As Integer)
			'Input some values to the cells
			cells(startRow, 1).PutValue(Integer.Parse(Me.dataTable1.Rows(index)("Quarter").ToString().Substring(5)))
			cells(startRow, 2).PutValue(CInt(Fix(Me.dataTable1.Rows(index)("Orders"))))
			cells(startRow, 3).PutValue(CDbl(CDec(Me.dataTable1.Rows(index)("Sales"))))
		End Sub
		Private Sub SetSalesByYearSubreportStyles(ByVal workbook As Workbook)
			'Create style and specify formatting attributes
			Dim styleIndex As Integer = workbook.Styles.Add()
			Dim style As Style = workbook.Styles(styleIndex)
			style.Font.IsBold = True
			style.Name = "Bold"
		End Sub
		Private Sub CreateSalesByYearSubreportHeader(ByVal workbook As Workbook, ByVal cells As Cells, ByVal startRow As Integer, ByVal year As String)
			'Input values and apply styles
			Dim style As Style = workbook.Styles("Bold")
			cells(startRow, 0).PutValue(year & " Summary")
			cells(startRow + 1, 1).PutValue("Quarter:")
			cells(startRow + 1, 2).PutValue("Orders Shipped:")
			cells(startRow + 1, 3).PutValue("Sales:")
			cells(startRow, 0).SetStyle(style)
			cells(startRow + 1, 1).SetStyle(style)
			cells(startRow + 1, 2).SetStyle(style)
			cells(startRow + 1, 3).SetStyle(style)

		End Sub

	End Class
End Namespace


