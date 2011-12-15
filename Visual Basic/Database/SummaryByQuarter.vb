Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for SummaryByQuarter.
	''' </summary>
	Public Class SummaryByQuarter
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)

		End Sub

		Public Function CreateSummaryByQuarter() As Workbook
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
				'Specify SQL and execute the query to fill the datatable
				Me.oleDbSelectCommand1.CommandText = "SELECT COUNT(Orders.OrderID) AS Orders, " & ControlChars.CrLf & "					SUM([Order Subtotals].Subtotal) AS Sales, FORMAT(Orders.ShippedDate, 'yyyy/Q') AS Quarter " & ControlChars.CrLf & "					FROM Orders INNER JOIN [Order Subtotals] ON Orders.OrderID = [Order Subtotals].OrderID " & ControlChars.CrLf & "					WHERE (Orders.ShippedDate IS NOT NULL) GROUP BY FORMAT(Orders.ShippedDate, 'yyyy/Q')"
				Me.oleDbDataAdapter1.Fill(Me.dataTable1)
			Catch
			Finally
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try
			'Get the worksheet
			Dim sheet As Worksheet = workbook.Worksheets("Sheet13")
			'Name the sheet
			sheet.Name = "Summary By Quarter"
			'Get the cells
			Dim cells As Cells = sheet.Cells

			'Create an arry of datatable with fields
			Dim quarterSummary(3) As DataTable
			For i As Integer = 0 To 3
				quarterSummary(i) = New DataTable()
				quarterSummary(i).Columns.Add("YearOrQuarter", GetType(Integer))
				quarterSummary(i).Columns.Add("Orders", GetType(Integer))
				quarterSummary(i).Columns.Add("Sales", GetType(Decimal))
			Next i

			'Adding some records to the datatables
			For i As Integer = 0 To Me.dataTable1.Rows.Count - 1
				Dim strQuarter As String = CStr(Me.dataTable1.Rows(i)("Quarter"))
				Dim quarter As Integer = Integer.Parse(strQuarter.Substring(strQuarter.Length - 1))
				Dim row As DataRow = quarterSummary(quarter - 1).NewRow()
				row("YearOrQuarter") = Integer.Parse(strQuarter.Substring(0, 4))
				row("Sales") = Me.dataTable1.Rows(i)("Sales")
				row("Orders") = Me.dataTable1.Rows(i)("Orders")
				quarterSummary(quarter - 1).Rows.Add(row)
			Next i

			'Replace some values in the workbook
			For i As Integer = 0 To 3
				workbook.Replace("&summary" & (i + 1).ToString(), quarterSummary(i))
			Next i
			'Remove the unnecessary worksheets in the workbook
            Dim iworkbook As Integer = 0
            Do While iworkbook < workbook.Worksheets.Count
                sheet = workbook.Worksheets(iworkbook)
                If sheet.Name <> "Summary By Quarter" Then
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


