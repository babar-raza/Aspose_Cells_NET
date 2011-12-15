Imports Microsoft.VisualBasic
Imports System

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for ProductsByCategory.
	''' </summary>
	Public Class ProductsByCategory
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)

		End Sub

		Public Function CreateProductsByCategory() As Workbook
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


			Me.dataTable1.Reset()
			Try
				'Specify an SQL and execute the query to fill the datatable
				Me.oleDbDataAdapter1.SelectCommand.CommandText = "SELECT Categories.CategoryName, Products.ProductName, Products.QuantityPerUnit, Products.UnitsInStock, Products.Discontinued, Categories.CategoryID, Products.ProductID FROM Categories INNER JOIN Products ON Categories.CategoryID = Products.CategoryID WHERE (Products.Discontinued <> Yes) ORDER BY Categories.CategoryName, Products.ProductName"
				Me.oleDbDataAdapter1.Fill(Me.dataTable1)
			Catch
			Finally
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try

			'Get a worksheet
			Dim sheet As Worksheet = workbook.Worksheets("Sheet7")
			'Name it
			sheet.Name = "Products By Category"
			'Get the cells
			Dim cells As Cells = sheet.Cells
			'Get the sheet vertical page breaks
			Dim vPageBreaks As VerticalPageBreakCollection = sheet.VerticalPageBreaks
			'Set row heights
			cells.SetRowHeight(4, 20.25)
			cells.SetRowHeight(5, 18.75)
			Dim currentRow As UShort = 4
			Dim currentColumn As Byte = 0

			Dim lastCategory As String = ""
			Dim thisCategory, nextCategory As String

			Dim productsCount As Integer = 0

			SetProductsByCategoryStyles(workbook)
			'Fill cells by inputing the values and apply styles to the data
			For i As Integer = 0 To Me.dataTable1.Rows.Count - 1
				thisCategory = CStr(Me.dataTable1.Rows(i)("CategoryName"))
				If thisCategory <> lastCategory Then
					currentRow = 4
					If i <> 0 Then
						currentColumn += 4
					End If
					CreateProductsByCategoryHeader(workbook, cells, currentRow, currentColumn, thisCategory)
					lastCategory = thisCategory
					currentRow += 2
				End If
				cells(currentRow, currentColumn).PutValue(CStr(Me.dataTable1.Rows(i)("ProductName")))
				cells(currentRow, CByte(currentColumn + 1)).PutValue(CShort(Fix(Me.dataTable1.Rows(i)("UnitsInStock"))))

				If i <> Me.dataTable1.Rows.Count - 1 Then
					nextCategory = CStr(Me.dataTable1.Rows(i + 1)("CategoryName"))
					If thisCategory <> nextCategory Then
						Dim style As Style = workbook.Styles("ProductsCount")
						cells(currentRow + 1, currentColumn).PutValue("Number of Products:")
						cells(currentRow + 1, currentColumn).SetStyle(style)

						style = workbook.Styles("CountNumber")
						cells(currentRow + 1, CByte(currentColumn + 1)).PutValue(productsCount + 1)
						cells(currentRow + 1, CByte(currentColumn + 1)).SetStyle(style)
						currentRow += 1
						productsCount = 0
						vPageBreaks.Add(0, currentColumn + 1)
					Else
						productsCount += 1
					End If
				Else
					Dim style As Style = workbook.Styles("ProductsCount")
					cells(currentRow + 1, currentColumn).PutValue("Number of Products:")
					cells(currentRow + 1, currentColumn).SetStyle(style)

					style = workbook.Styles("CountNumber")
					cells(currentRow + 1, CByte(currentColumn + 1)).PutValue(productsCount + 1)
					cells(currentRow + 1, CByte(currentColumn + 1)).SetStyle(style)
				End If
				currentRow += 1
			Next i

			'Remove the unnecessary worksheets in the workbook
            Dim iworkbook As Integer = 0
            Do While iworkbook < workbook.Worksheets.Count
                sheet = workbook.Worksheets(iworkbook)
                If sheet.Name <> "Products By Category" Then
                    workbook.Worksheets.RemoveAt(iworkbook)
                    iworkbook -= 1
                End If
                iworkbook += 1
            Loop
			'Get the generated workbook
			Return workbook
		End Function

		Private Sub SetProductsByCategoryStyles(ByVal workbook As Workbook)
			'Create a style with some specific formatting attributes
			Dim styleIndex As Integer = workbook.Styles.Add()
			Dim style As Style = workbook.Styles(styleIndex)
			style.Font.IsItalic = True
			style.Font.IsBold = True
			style.Font.Size = 16
			style.HorizontalAlignment = TextAlignmentType.Right
			style.Name = "Category"

			'Create a style with some specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Font.Size = 16
			style.Font.IsBold = True
			style.HorizontalAlignment = TextAlignmentType.Left
			style.Name = "CategoryName"

			'Create a style with some specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Font.Size = 14
			style.Font.IsBold = True
			style.Font.IsItalic = True
			style.HorizontalAlignment = TextAlignmentType.Left
			style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Medium
			style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Medium
			style.Name = "ProductName"

			'Create a style with some specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Font.Size = 14
			style.Font.IsBold = True
			style.Font.IsItalic = True
			style.HorizontalAlignment = TextAlignmentType.Right
			style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Medium
			style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Medium
			style.Name = "UnitsInStock"

			'Create a style with some specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Font.IsBold = True
			style.Font.IsItalic = True
			style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style.Name = "ProductsCount"

			'Create a style with some specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.HorizontalAlignment = TextAlignmentType.Left
			style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style.Name = "CountNumber"

		End Sub
		Private Sub CreateProductsByCategoryHeader(ByVal workbook As Workbook, ByVal cells As Cells, ByVal startRow As UShort, ByVal startColumn As Byte, ByVal categoryName As String)
			'Input values and apply the styles to the cells

			Dim style As Style = workbook.Styles("Category")
			cells(startRow, startColumn).PutValue("Category:")
			cells(startRow, startColumn).SetStyle(style)

			style = workbook.Styles("CategoryName")
			cells(startRow, CByte(startColumn + 1)).PutValue(categoryName)
			cells(startRow, CByte(startColumn + 1)).SetStyle(style)

			style = workbook.Styles("ProductName")
			cells(startRow + 1, startColumn).PutValue("Product Name")
			cells(startRow + 1, startColumn).SetStyle(style)

			style = workbook.Styles("UnitsInStock")
			cells(startRow + 1, CByte(startColumn + 1)).PutValue("Units In Stock:")
			cells(startRow + 1, CByte(startColumn + 1)).SetStyle(style)
		End Sub


	End Class
End Namespace
