Imports Microsoft.VisualBasic
Imports System
Imports System.Data

Namespace Aspose.Cells.Demos
    ''' <summary>
    ''' Summary description for SummaryByYear.
    ''' </summary>
    Public Class SummaryByYear
        Inherits DbBase
        Public Sub New(ByVal path As String)
            MyBase.New(path)
        End Sub

        Public Function CreateSummaryByYear() As Workbook
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
                Me.oleDbSelectCommand1.CommandText = "SELECT COUNT(Orders.OrderID) AS Orders, SUM([Order Subtotals].Subtotal) AS Sales, " & ControlChars.CrLf & "					FORMAT(Orders.ShippedDate, 'yyyy/Q') AS Quarter FROM Orders INNER JOIN [Order Subtotals] " & ControlChars.CrLf & "					ON Orders.OrderID = [Order Subtotals].OrderID WHERE (Orders.ShippedDate IS NOT NULL) " & ControlChars.CrLf & "					GROUP BY FORMAT(Orders.ShippedDate, 'yyyy/Q')"
                Me.oleDbDataAdapter1.Fill(Me.dataTable1)
            Catch
            Finally
                If Me.oleDbConnection1 IsNot Nothing Then
                    Me.oleDbConnection1.Close()
                End If
            End Try

            'Get the worksheet
            Dim sheet As Worksheet = workbook.Worksheets("Sheet14")
            'Name the sheet
            sheet.Name = "Summary By Year"
            'Get the cells collection
            Dim cells As Cells = sheet.Cells
            'Create an array of datatables with specific fields
            Dim yearSummary(2) As DataTable
            For i As Integer = 0 To 2
                yearSummary(i) = New DataTable()
                yearSummary(i).Columns.Add("YearOrQuarter", GetType(Integer))
                yearSummary(i).Columns.Add("Orders", GetType(Integer))
                yearSummary(i).Columns.Add("Sales", GetType(Decimal))
            Next i
            'Adding records to the datatables
            For i As Integer = 0 To Me.dataTable1.Rows.Count - 1
                Dim strQuarter As String = CStr(Me.dataTable1.Rows(i)("Quarter"))
                Dim year As Integer = Integer.Parse(strQuarter.Substring(0, 4)) - 1994
                Dim row As DataRow = yearSummary(year).NewRow()
                row("YearOrQuarter") = Integer.Parse(strQuarter.Substring(strQuarter.Length - 1))
                row("Sales") = Me.dataTable1.Rows(i)("Sales")
                row("Orders") = Me.dataTable1.Rows(i)("Orders")
                yearSummary(year).Rows.Add(row)
            Next i
            'Replace some values in the workbook
            For i As Integer = 0 To 2
                workbook.Replace("&summary" & (i + 1).ToString(), yearSummary(i))
            Next i
            'Remove the unnecessary worksheets in the workbook
            Dim iworkbook As Integer = 0
            Do While iworkbook < workbook.Worksheets.Count
                sheet = workbook.Worksheets(iworkbook)
                If sheet.Name <> "Summary By Year" Then
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

