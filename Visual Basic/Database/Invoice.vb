Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Drawing
Imports Aspose.Cells.Drawing
Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for Invoice.
	''' </summary>
	Public Class Invoice
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)

		End Sub

		Public Function CreateInvoice() As Workbook
			Try
				DBInit()

				'Specify SQL for command text
				Me.oleDbDataAdapter1.SelectCommand.CommandText = "SELECT DISTINCTROW OrderID FROM Orders ORDER BY OrderID DESC"
				'Fill the datatable 
				Me.oleDbDataAdapter1.Fill(Me.dataTable1)
			Catch
			Finally
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try

			Dim dtInvoice(Me.dataTable1.Rows.Count - 1) As DataTable

			'for(int i = 0; i < dtInvoice.Length; i ++)
			'We generate invoices for the first 100 orders for demo only. If you want to
			'generate all invoices,uncomment the line above and comment the line below.
			For i As Integer = 0 To 49
				dtInvoice(i) = Me.ReadInvoice(Me.dataTable1.Rows(i)(0).ToString())
			Next i

			'Create the workbook
			Dim workbook As New Workbook()
			'Get all the worksheets
			Dim sheets As WorksheetCollection = workbook.Worksheets
			'get the first worksheet
			Dim sheet As Worksheet = sheets(0)
			'Name the worksheet
			sheet.Name = "Invoice"
			'Get the sheet cells
			Dim cells As Cells = sheet.Cells
			Dim startRow As Integer = 0

			SetInvoiceStyles(workbook)
			Dim imagePath As String = path & "\Image"
			'for(int i = 0; i < dtInvoice.Length; i ++)
			'We generate invoices for the first 100 orders for demo only. If you want to
			'generate all invoices,uncomment the line above and comment the line below.
			For i As Integer = 0 To 49
				'Add picture(s)
				sheet.Pictures.Add(startRow, 0, startRow + 2, 1, imagePath & "\logo.jpg")
				Dim picIndex As Integer = sheet.Pictures.Add(startRow, 1, startRow + 2, 2, imagePath & "\namelogo.jpg")
				Dim pic As Picture = sheet.Pictures(picIndex)
				pic.UpperDeltaY = 100

				CreateInvoiceHeader(cells, workbook, dtInvoice(i), startRow)
				startRow += 11
				CreateOrder(cells, workbook, dtInvoice(i), startRow, Me.dataTable1.Rows(i)(0).ToString())
				startRow += 4
				CreateOrderDetail(cells, workbook, dtInvoice(i), startRow)

				startRow += dtInvoice(i).Rows.Count + 1
				'Add horizontal page break(s)
				sheet.HorizontalPageBreaks.Add(startRow - 1, 0)
			Next i

			'Get the workbook (generated)
			Return workbook

		End Function
		Private Function ReadInvoice(ByVal orderID As String) As DataTable
			Try
				'Specify SQL
				Dim invoiceQuery As String = "SELECT DISTINCTROW Invoices.* FROM Invoices WHERE Invoices.OrderID=" & orderID
				'Specify the command
				Me.oleDbDataAdapter2.SelectCommand.CommandText = invoiceQuery
			Catch
			Finally
				If Me.oleDbConnection1 IsNot Nothing Then
					Me.oleDbConnection1.Close()
				End If
			End Try
			'Create a datatable and fill it with data based on the query
			Dim dtInvoice As New DataTable()
			Me.oleDbDataAdapter2.Fill(dtInvoice)
			'Retrieve the datatable
			Return dtInvoice
		End Function

		Private Sub SetInvoiceStyles(ByVal workbook As Workbook)
			'Add LightBlue and DarkBlue colors to color palette
			workbook.ChangePalette(Color.LightBlue, 54)
			workbook.ChangePalette(Color.DarkBlue, 55)

			'Create a style with specific formatting attributes
			Dim style As Style
			Dim styleIndex As Integer = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Font.Size = 12
			style.Font.IsBold = True
			style.Font.Color = Color.White
			style.ForegroundColor = Color.LightBlue
			style.Pattern = BackgroundType.Solid
			style.HorizontalAlignment = TextAlignmentType.Center
			style.Name = "Font12Center"

			'Create a style with specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Font.Size = 12
			style.Font.IsBold = True
			style.Font.Color = Color.White
			style.ForegroundColor = Color.LightBlue
			style.Pattern = BackgroundType.Solid
			style.HorizontalAlignment = TextAlignmentType.Left
			style.Name = "Font12Left"

			'Create a style with specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Font.Size = 12
			style.Font.IsBold = True
			style.Font.Color = Color.White
			style.ForegroundColor = Color.LightBlue
			style.Pattern = BackgroundType.Solid
			style.HorizontalAlignment = TextAlignmentType.Right
			style.Name = "Font12Right"

			'Create a style with specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Number = 7
			style.Name = "Number7"

			'Create a style with specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Number = 9
			style.Name = "Number9"

			'Create a style with specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.HorizontalAlignment = TextAlignmentType.Center
			style.Name = "Center"

			'Create a style with specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Font.Size = 16
			style.Font.IsBold = True
			style.Font.Color = Color.DarkBlue
			style.Name = "Darkblue"

			'Create a style with specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Font.Size = 12
			style.Font.IsBold = True
			style.Font.Color = Color.DarkBlue
			style.Name = "Darkblue12"

			'Create a style with specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Font.IsItalic = True
			style.Font.Color = Color.DarkBlue
			style.Name = "DarkblueItalic"

			'Create a style with specific formatting attributes
			styleIndex = workbook.Styles.Add()
			style = workbook.Styles(styleIndex)
			style.Borders(BorderType.BottomBorder).Color = Color.Black
			style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Medium
			style.Name = "BlackMedium"


		End Sub

		Private Sub CreateOrderDetail(ByVal cells As Cells, ByVal workbook As Workbook, ByVal dtInvoice As DataTable, ByVal startRow As Integer)
			'Define some styles
			Dim style1, style2, style3 As Style
			'Get styles
			style1 = workbook.Styles("Number7")
			style2 = workbook.Styles("Number9")
			style3 = workbook.Styles("Center")

			'Fill cells based on differnt datatable fields
			'Apply styles to cells too
			For i As Integer = 0 To dtInvoice.Rows.Count - 1
				cells(startRow + i, 0).PutValue(CInt(Fix(dtInvoice.Rows(i)("ProductID"))))
				cells(startRow + i, 0).SetStyle(style3)
				cells(startRow + i, 1).PutValue(CStr(dtInvoice.Rows(i)("ProductName")))
				cells(startRow + i, 3).PutValue(CShort(Fix(dtInvoice.Rows(i)("Quantity"))))
				cells(startRow + i, 4).PutValue(CDbl(CDec(dtInvoice.Rows(i)("UnitPrice"))))
				cells(startRow + i, 4).SetStyle(style1)
				cells(startRow + i, 5).PutValue(CSng(dtInvoice.Rows(i)("Discount")))
				cells(startRow + i, 5).SetStyle(style2)
				cells(startRow + i, 6).PutValue(CDbl(CDec(dtInvoice.Rows(i)("ExtendedPrice"))))
				cells(startRow + i, 6).SetStyle(style1)
			Next i
		End Sub

		Private Sub CreateOrder(ByVal cells As Cells, ByVal workbook As Workbook, ByVal dtInvoice As DataTable, ByVal startRow As Integer, ByVal orderID As String)
			'Set row heights for some rows
			cells.SetRowHeight(startRow, 14)
			cells.SetRowHeight(startRow + 3, 14)

			'Set column widths for some columns
			cells.SetColumnWidth(1, 16)
			cells.SetColumnWidth(2, 16)
			cells.SetColumnWidth(3, 16)
			cells.SetColumnWidth(4, 16)
			cells.SetColumnWidth(5, 16)
			cells.SetColumnWidth(6, 18)
			'Get the style
			Dim style As Style = workbook.Styles("Font12Center")
			'Apply the style to the cells
			For i As Byte = 0 To 6
				cells(startRow, i).SetStyle(style)
				cells(startRow + 3, i).SetStyle(style)
			Next i
			'Get the style
			style = workbook.Styles("Center")
			'Apply the style to the cells
			For i As Byte = 0 To 6
				cells(startRow + 1, i).SetStyle(style)
			Next i

			'Input values to the cells based on the datatable
			cells(startRow, 0).PutValue("Order ID:")
			cells(startRow + 1, 0).PutValue(Integer.Parse(orderID))
			cells(startRow, 1).PutValue("Customer ID:")
			cells(startRow + 1, 1).PutValue(CStr(dtInvoice.Rows(0)("CustomerID")))
			cells(startRow, 2).PutValue("Salesperson:")
			cells(startRow + 1, 2).PutValue(CStr(dtInvoice.Rows(0)("Salesperson")))
			cells(startRow, 3).PutValue("Order Date:")
			cells(startRow + 1, 3).PutValue((CDate(dtInvoice.Rows(0)("OrderDate"))).ToString("D"))
			cells(startRow, 4).PutValue("Required Date:")
			cells(startRow + 1, 4).PutValue((CDate(dtInvoice.Rows(0)("RequiredDate"))).ToString("D"))
			cells(startRow, 5).PutValue("Shipped Date:")
			If dtInvoice.Rows(0)("ShippedDate") IsNot DBNull.Value Then
				cells(startRow + 1, 5).PutValue((CDate(dtInvoice.Rows(0)("ShippedDate"))).ToString("D"))
			End If
			cells(startRow, 6).PutValue("Ship Via:")
			cells(startRow + 1, 6).PutValue(CStr(dtInvoice.Rows(0)("Shippers.CompanyName")))

			cells(startRow + 3, 0).PutValue("Product ID:")
			cells(startRow + 3, 1).PutValue("Product")
			cells(startRow + 3, 2).PutValue(" Name:")
			cells(startRow + 3, 3).PutValue("Quantity:")
			cells(startRow + 3, 4).PutValue("Unit Price:")
			cells(startRow + 3, 5).PutValue("Discount:")
			cells(startRow + 3, 6).PutValue("Extended Price:")

			'Get the style and apply it to the cell(s)
			style = workbook.Styles("Font12Right")
			cells(startRow + 3, 1).SetStyle(style)

			'Get the style and apply it to the cell(s)
			style = workbook.Styles("Font12Left")
			cells(startRow + 3, 2).SetStyle(style)
		End Sub

		Private Sub CreateInvoiceHeader(ByVal cells As Cells, ByVal workbook As Workbook, ByVal dtInvoice As DataTable, ByVal startRow As Integer)

			'Set row height and column width 
			cells.SetRowHeight(startRow, 24)
			cells.SetColumnWidth(0, 12)

			'Input a value and set its style
			cells(startRow, 5).PutValue("INVOICE")
			Dim style As Style = workbook.Styles("Darkblue")
			cells(startRow, 5).SetStyle(style)

			'Get the style and apply it to the cells
			style = workbook.Styles("BlackMedium")
			For i As Integer = 0 To Byte.MaxValue - 1
				cells(startRow + 2, CByte(i)).SetStyle(style)
			Next i

			'Input some values to the cells
			cells(startRow + 3, 0).PutValue("One Portals Way, Twin Points WA 98156")
			cells(startRow + 4, 0).PutValue("Phone:1-206-555-1417 Fax:1-206")
			style = workbook.Styles("DarkblueItalic")
			cells(startRow + 3, 0).SetStyle(style)
			cells(startRow + 4, 0).SetStyle(style)

			'Get the current date
			Dim currentDate As DateTime = DateTime.Today
			Dim strTime As String = currentDate.ToString("D")
			'Input date
			cells(startRow + 3, 5).PutValue("Date:")
			cells(startRow + 3, 6).PutValue(strTime)

			'Input a value
			cells(startRow + 6, 0).PutValue("Ship To:")
			'Get the style
			style = workbook.Styles("Darkblue12")
			'Apply the style to a cell
			cells(startRow + 6, 0).SetStyle(style)
			'Set the related row height
			cells.SetRowHeight(startRow + 6, 16)
			'Input a value and apply style to it
			cells(startRow + 6, 4).PutValue("Bill To:")
			cells(startRow + 6, 4).SetStyle(style)
			'Apply the style to a cell
			cells(startRow + 3, 5).SetStyle(style)
			'Input values
			If dtInvoice.Rows(0)(0) IsNot DBNull.Value Then
				cells(startRow + 6, 1).PutValue(CStr(dtInvoice.Rows(0)(0)))
				cells(startRow + 6, 5).PutValue(CStr(dtInvoice.Rows(0)(0)))
			End If
			If dtInvoice.Rows(0)(1) IsNot DBNull.Value Then
				cells(startRow + 7, 1).PutValue(CStr(dtInvoice.Rows(0)(1)))
				cells(startRow + 7, 5).PutValue(CStr(dtInvoice.Rows(0)(1)))
			End If

			Dim strDest As String = ""
			If dtInvoice.Rows(0)(2) IsNot DBNull.Value Then
				strDest &= dtInvoice.Rows(0)(2)
			End If

			If dtInvoice.Rows(0)(3) IsNot DBNull.Value Then
				strDest &= " " & dtInvoice.Rows(0)(3)
			End If

			If dtInvoice.Rows(0)(4) IsNot DBNull.Value Then
				strDest &= " " & dtInvoice.Rows(0)(4)
			End If

			strDest.TrimStart(" "c)

			If strDest <> "" Then
				cells(startRow + 8, 1).PutValue(strDest)
				cells(startRow + 8, 5).PutValue(strDest)
			End If
			cells(startRow + 9, 1).PutValue(CStr(dtInvoice.Rows(0)(5)))
			cells(startRow + 9, 5).PutValue(CStr(dtInvoice.Rows(0)(5)))

		End Sub

	End Class
End Namespace


