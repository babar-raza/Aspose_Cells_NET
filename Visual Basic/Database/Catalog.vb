Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for Catalog.
	''' </summary>
	Public Class Catalog
		Inherits DbBase
		Public Sub New(ByVal path As String)
			MyBase.New(path)

		End Sub

		Public Function CreateCatalog() As Workbook

			Try
				DBInit()
			Catch
			End Try

			'Open a template file
		Dim designerFile As String = MapPath("~/Designer/Northwind.xls")
		Dim workbook As New Workbook(designerFile)


			ReadCategory()
			'Create a new datatable
			Dim dataTable2 As New DataTable()
			'Get a worksheet
			Dim sheet As Worksheet = workbook.Worksheets("Sheet2")
			'Name the sheet
			sheet.Name = "Catalog"
			'Get the worksheet cells
			Dim cells As Cells = sheet.Cells

			Dim currentRow As Integer = 55

			'Add LightGray color to color palette
			workbook.ChangePalette(Color.LightGray, 55)
			'Get the workbook's styles collection
			Dim styles As StyleCollection = workbook.Styles
			'Set CategoryName style with formatting attributes
			Dim styleIndex As Integer = styles.Add()
			Dim styleCategoryName As Style = styles(styleIndex)
			styleCategoryName.Font.Size = 14
			styleCategoryName.Font.Color = Color.Blue
			styleCategoryName.Font.IsBold = True
			styleCategoryName.Font.Name = "Times New Roman"

			'Set Description style with formatting attributes
			styleIndex = styles.Add()
			Dim styleDescription As Style = styles(styleIndex)
			styleDescription.Font.Name = "Times New Roman"
			styleDescription.Font.Color = Color.Blue
			styleDescription.Font.IsItalic = True

			'Set ProductName style with formatting attributes
			styleIndex = styles.Add()
			Dim styleProductName As Style = styles(styleIndex)
			styleProductName.Font.IsBold = True

			'Set Title style with formatting attributes
			styleIndex = styles.Add()
			Dim styleTitle As Style = styles(styleIndex)
			styleTitle.Font.IsBold = True
			styleTitle.Font.IsItalic = True
			styleTitle.ForegroundColor = Color.LightGray

			styleIndex = styles.Add()
			Dim styleNumber As Style = styles(styleIndex)
			styleNumber.Font.Name = "Times New Roman"
			styleNumber.Number = 8

			'Create the styleflag struct
			Dim styleflag As New StyleFlag()
			styleflag.All = True
			'Get the horizontal page breaks collection
			Dim hPageBreaks As HorizontalPageBreakCollection = sheet.HorizontalPageBreaks

			'Specify SQL for the command
			Dim cmd As String = "SELECT ProductName, ProductID, QuantityPerUnit, " & "UnitPrice FROM Products"
			For i As Integer = 0 To Me.dataTable1.Rows.Count - 1
				currentRow += 2
				cells.SetRowHeight(currentRow, 20)
				cells(currentRow, 1).SetStyle(styleCategoryName)
				Dim categoriesRow As DataRow = Me.dataTable1.Rows(i)

				'Write CategoryName
				cells(currentRow, 1).PutValue(CStr(categoriesRow("CategoryName")))

				'Write Description
				currentRow += 1
				cells(currentRow, 1).PutValue(CStr(categoriesRow("Description")))
				cells(currentRow, 1).SetStyle(styleDescription)

				dataTable2.Clear()

				'Execuate command and fill the datatable
				Try
					Me.oleDbDataAdapter2 = New OleDbDataAdapter()
					Dim cmdText As String = cmd & " where categoryid = " & categoriesRow("CategoryID").ToString()
					Me.oleDbDataAdapter2.SelectCommand = New OleDbCommand(cmdText, Me.oleDbConnection1)
					Me.oleDbConnection1.Open()
					oleDbDataAdapter2.Fill(dataTable2)
				Catch
				Finally
					oleDbDataAdapter2.Dispose()
					Me.oleDbConnection1.Close()
				End Try

				currentRow += 2
				'Import the datatable to the sheet
				cells.ImportDataTable(dataTable2, True, currentRow, 1)
				'Create a range
				Dim range As Range = cells.CreateRange(currentRow, 1, 1, 4)
				'Apply style to the range
				range.ApplyStyle(styleTitle, styleflag)
				'Create a range
				range = cells.CreateRange(currentRow + 1, 1, dataTable2.Rows.Count, 1)
				'Apply style to the range
				range.ApplyStyle(styleProductName, styleflag)
				'Create a range
				range = cells.CreateRange(currentRow + 1, 4, dataTable2.Rows.Count, 1)
				'Apply style to the range
				range.ApplyStyle(styleNumber, styleflag)

				currentRow += dataTable2.Rows.Count
				'Apply horizontal page breaks
				hPageBreaks.Add(currentRow, 0)
			Next i

			'Remove the unnecessary worksheets in the workbook
            Dim iworkbook As Integer = 0
            Do While iworkbook < workbook.Worksheets.Count
                sheet = workbook.Worksheets(iworkbook)
                If sheet.Name <> "Catalog" Then
                    workbook.Worksheets.RemoveAt(iworkbook)
                    iworkbook -= 1
                End If

                iworkbook += 1
            Loop
			'Return the generated workbook
			Return workbook
		End Function

		Private Sub ReadCategory()
			'Execute the command and fill a datatable
			Try
				Me.oleDbConnection1.Open()
				Me.oleDbSelectCommand1.CommandText = "SELECT CategoryID, CategoryName, Description FROM Categories"
				Me.oleDbDataAdapter1.Fill(Me.dataTable1)
			Catch
			Finally
				Me.oleDbDataAdapter1.Dispose()
				Me.oleDbConnection1.Close()
			End Try

		End Sub

	End Class
End Namespace


