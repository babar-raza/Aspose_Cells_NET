Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports Aspose.Cells
Imports Aspose.Cells.Drawing
Imports Aspose.Cells.Charts

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for SalesByCategory.
	''' </summary>
	Public Class SalesByCategory
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)
			'
			' TODO: Add constructor logic here
			'
		End Sub

		Public Function CreateSalesByCategory() As Workbook
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
				Me.oleDbDataAdapter1.SelectCommand.CommandText = "SELECT DISTINCTROW Categories.CategoryID, " & ControlChars.CrLf & "					Categories.CategoryName, Products.ProductName, SUM([Order Details Extended].ExtendedPrice) AS ProductSales" & ControlChars.CrLf & "				FROM  Categories " & ControlChars.CrLf & "				INNER JOIN" & ControlChars.CrLf & "					(Products INNER JOIN (Orders INNER JOIN [Order Details Extended] ON" & ControlChars.CrLf & "					Orders.OrderID = [Order Details Extended].OrderID) ON Products.ProductID = [Order Details Extended].ProductID) ON Categories.CategoryID = Products.CategoryID" & ControlChars.CrLf & "				WHERE" & ControlChars.CrLf & "					(((Orders.OrderDate) BETWEEN #1/1/1995# AND " & ControlChars.CrLf & "						#12/31/1995#)) GROUP BY Categories.CategoryID ,  Categories.CategoryName ,  Products.ProductName ORDER BY Categories.CategoryName"
				Me.oleDbDataAdapter1.Fill(Me.dataTable1)
			Catch
			Finally
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try

			'Get the worksheet
			Dim sheet As Worksheet = workbook.Worksheets("Sheet8")
			'Name the worksheet
			sheet.Name = "Sales By Category"
			'Get the cells
			Dim cells As Cells = sheet.Cells
			'Get the vertical page breaks
			Dim vPageBreaks As VerticalPageBreakCollection = sheet.VerticalPageBreaks
			Dim currentRow As Integer = 2
			Dim currentColumn As Byte = 0

			Dim lastCategory As String = ""
			Dim thisCategory, nextCategory As String

			SetSalesByCategoryStyles(workbook)
			'Fill cells with source data and apply styles
			For i As Integer = 0 To Me.dataTable1.Rows.Count - 1
				thisCategory = CStr(Me.dataTable1.Rows(i)("CategoryName"))
				If thisCategory <> lastCategory Then
					currentRow = 2
					If i <> 0 Then
						currentColumn += 15
					End If
					CreateSalesByCategoryHeader(workbook, cells, currentRow, currentColumn, thisCategory)
					lastCategory = thisCategory
					currentRow += 2
				End If
				cells(currentRow, currentColumn).PutValue(CStr(Me.dataTable1.Rows(i)("ProductName")))
				cells(currentRow, CByte(currentColumn + 1)).PutValue(CDbl(CDec(Me.dataTable1.Rows(i)("ProductSales"))))

				cells(currentRow, CByte(currentColumn + 1)).SetStyle(workbook.Styles("Sales"))

				cells.SetColumnWidth(currentColumn, 27)
				cells.SetColumnWidth(CByte(currentColumn + 1), 15)

				If i <> Me.dataTable1.Rows.Count - 1 Then
					nextCategory = CStr(Me.dataTable1.Rows(i + 1)("CategoryName"))
					If thisCategory <> nextCategory Then
						vPageBreaks.Add(0, currentColumn + 1)
						CreateChart(workbook, sheet, currentRow, currentColumn)
					End If
				Else
					CreateChart(workbook, sheet, currentRow, currentColumn)
				End If
				currentRow += 1
			Next i
			'Remove the unnecessary worksheets
            Dim iworkbook As Integer = 0
            Do While iworkbook < workbook.Worksheets.Count
                sheet = workbook.Worksheets(iworkbook)
                If sheet.Name <> "Sales By Category" Then
                    workbook.Worksheets.RemoveAt(iworkbook)
                    iworkbook -= 1
                End If
                iworkbook += 1
            Loop
			'Get the workbook (generated)
			Return workbook
		End Function

		Private Sub CreateChart(ByVal workbook As Workbook, ByVal sheet As Worksheet, ByVal currentRow As Integer, ByVal currentColumn As Integer)
			'Add a bar chart
			Dim chartIndex As Integer = sheet.Charts.Add(ChartType.Bar, 4, currentColumn + 3, 26, currentColumn + 14)
			'Get the chart
			Dim chart As Chart = sheet.Charts(chartIndex)
			'Make the legends invisible
			chart.ShowLegend = False
			Dim startCell As String = CellsHelper.CellIndexToName(4, currentColumn + 1)
			Dim endCell As String = CellsHelper.CellIndexToName(currentRow, currentColumn + 1)
			'Set the nseries
			chart.NSeries.Add(startCell & ":" & endCell, True)
			'Set the fill format for the plot area
			Dim fillFormat As FillFormat = chart.PlotArea.Area.FillFormat
			fillFormat.SetPresetColorGradient(GradientPresetType.Daybreak, GradientStyleType.Vertical, 1)
			'Set the category data
			startCell = CellsHelper.CellIndexToName(4, currentColumn)
			endCell = CellsHelper.CellIndexToName(currentRow, currentColumn)
			chart.NSeries.CategoryData = startCell & ":" & endCell
		End Sub
		Private Sub CreateSalesByCategoryHeader(ByVal workbook As Workbook, ByVal cells As Cells, ByVal currentRow As Integer, ByVal currentColumn As Byte, ByVal categoryName As String)
			'Input data and apply style
			cells(currentRow, currentColumn).PutValue(categoryName)
			cells(currentRow, currentColumn).SetStyle(workbook.Styles("Header"))
		End Sub

		Private Sub SetSalesByCategoryStyles(ByVal workbook As Workbook)
			'Create a style with specific formatting attributes
			Dim styleIndex As Integer = workbook.Styles.Add()
			Dim style As Style = workbook.Styles(styleIndex)
			style.Number = 7
			style.Name = "Sales"

			'Create a style withe specific set of attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Font.Size = 14
			style.Font.IsBold = True
			style.Font.IsItalic = True
			style.Font.Color = Color.Yellow
			style.ForegroundColor = Color.Blue
			style.Pattern = BackgroundType.Solid
			style.Name = "Header"
		End Sub

	End Class
End Namespace


