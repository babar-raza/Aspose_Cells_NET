Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Web
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports Aspose.Cells
Imports System.Data.OleDb

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for ImportingData.
	''' </summary>
	Public Class ImportingData
		Inherits System.Web.UI.Page
		Protected ImportingDataType As System.Web.UI.WebControls.DropDownList
		Protected WithEvents btnCreateReport As System.Web.UI.WebControls.Button
		Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

		Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			' Put user code to initialize the page here
		End Sub

		#Region "Web Form Designer generated code"
		Overrides Protected Sub OnInit(ByVal e As EventArgs)
			'
			' CODEGEN: This call is required by the ASP.NET Web Form Designer.
			'
			InitializeComponent()
			MyBase.OnInit(e)
		End Sub

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
'			Me.btnCreateReport.Click += New System.EventHandler(Me.btnCreateReport_Click);
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region

		Private Sub btnCreateReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreateReport.Click
			'Instantiate a new workbook
			Dim workbook As New Workbook()

			'Get the first worksheet in the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'switch case for dropdown's selected value
			Select Case ImportingDataType.SelectedItem.Text
				'if selected text is "Array"
				Case "Array"
					ImportArray(sheet)
				'if selected text is "ArrayList"
				Case "ArrayList"
					ImportArrayList(sheet)
				'if selected text is "DataColumn"
				Case "DataColumn"
					ImportDataColumn(sheet)
				'if selected text is "DataGrid"
				Case "DataGrid"
					ImportDataGrid(sheet)
				'if selected text is "DataTable"
				Case "DataTable"
					ImportDataTable(sheet)
				'if selected text is "DataView"
				Case "DataView"
					ImportDataView(sheet)
				'if selected text is "FormulaArray"
				Case "FormulaArray"
					ImportFormulaArray(sheet)
				'if selected text is "FromDataReader"
				Case "FromDataReader"
					ImportFromDataReader(sheet)
				'if selected text is "ObjectArray"
				Case "ObjectArray"
					ImportObjectArray(sheet)
				'if selected text is "TwoDimensionArray"
				Case "TwoDimensionArray"
					ImportTwoDimensionArray(sheet)
				Case Else
					Return
			End Select

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "ImportingData.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "ImportingData.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()

		End Sub

		Private Sub ImportArray(ByVal sheet As Worksheet)
			'Get the cells collection in the worksheet
			Dim cells As Cells = sheet.Cells

			'Put a string value into a cell
			cells("A1").PutValue("Import an Array")

			'Get Style Object 
			Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

			'the font text is set to bold
			style.Font.IsBold = True

			'Apply style to the cell 
			cells("A1").SetStyle(style)

			'Create a string array of values
			Dim names() As String = { "Tom", "John", "Kelly" }

			'Import the array to the sheet cells
			sheet.Cells.ImportArray(names, 1, 0, True)
		End Sub

		Private Sub ImportArrayList(ByVal sheet As Worksheet)
			'Get the cells collection in the worksheet
			Dim cells As Cells = sheet.Cells

			'Put a string value into a cell
			cells("A1").PutValue("Import an ArrayList")

			'Get Style Object 
			Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

			'the font text is set to bold
			style.Font.IsBold = True

			'Apply style to the cell 
			cells("A1").SetStyle(style)

			'Create an arraylist and fill some values to it
			Dim list As New ArrayList()
			list.Add("Tom")
			list.Add("John")
			list.Add("Kelly")

			'Import the arraylist to the sheet cells
			sheet.Cells.ImportArrayList(list, 1, 0, True)
		End Sub

		Private Sub ImportDataColumn(ByVal sheet As Worksheet)
			'Get the cells collection in the worksheet
			Dim cells As Cells = sheet.Cells

			'Put a string value to a cell
			cells("A1").PutValue("Import a DataColumn")

			'Get Style Object 
			Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

			'the font text is set to bold
			style.Font.IsBold = True

			'Apply style to the cell 
			cells("A1").SetStyle(style)

			'Create a datatable and add three columns to it
			Dim dataTable As New DataTable("Products")
			dataTable.Columns.Add("Product ID", GetType(Int32))
			dataTable.Columns.Add("Product Name", GetType(String))
			dataTable.Columns.Add("Units In Stock", GetType(Int32))

			'Add the first record to it
			Dim dr As DataRow = dataTable.NewRow()
			dr(0) = 1
			dr(1) = "Aniseed Syrup"
			dr(2) = 15
			dataTable.Rows.Add(dr)

			'Add a second record to it
			dr = dataTable.NewRow()
			dr(0) = 2
			dr(1) = "Boston Crab Meat"
			dr(2) = 123
			dataTable.Rows.Add(dr)

			'Import the datacolumn in the datatable to the sheet cells
			sheet.Cells.ImportDataColumn(dataTable, True, 1, 0, 1, False)
		End Sub

		Private Sub ImportDataGrid(ByVal sheet As Worksheet)
			'Get the cells collection in the worksheet
			Dim cells As Cells = sheet.Cells

			'Put a string value into a cell
			sheet.Cells("A1").PutValue("Import a DataGrid")

			'Get Style Object 
			Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

			'the font text is set to bold
			style.Font.IsBold = True

			'Apply style to the cell 
			cells("A1").SetStyle(style)

			'Create a datatable and add three columns to it
			Dim dataTable As New DataTable("Products")
			dataTable.Columns.Add("Product ID", GetType(Int32))
			dataTable.Columns.Add("Product Name", GetType(String))
			dataTable.Columns.Add("Units In Stock", GetType(Int32))

			'Add the first record to it
			Dim dr As DataRow = dataTable.NewRow()
			dr(0) = 1
			dr(1) = "Aniseed Syrup"
			dr(2) = 15
			dataTable.Rows.Add(dr)

			'Add the second record to it
			dr = dataTable.NewRow()
			dr(0) = 2
			dr(1) = "Boston Crab Meat"
			dr(2) = 123
			dataTable.Rows.Add(dr)

			'Create a datagrid
			Dim dataGrid As New DataGrid()

			'set its datasource
			dataGrid.DataSource = dataTable

			'bind data
			dataGrid.DataBind()

			'Import the datagrid to sheet cells
			sheet.Cells.ImportDataGrid(dataGrid, 1, 0, False)

			'Autofit all the columns in the sheet
			sheet.AutoFitColumns()
		End Sub

		Private Sub ImportDataTable(ByVal sheet As Worksheet)
			'Get the cells collection in the worksheet
			Dim cells As Cells = sheet.Cells

			'Put a string value into a cell
			sheet.Cells("A1").PutValue("Import a DataTable")

			'Get Style Object 
			Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

			'the font text is set to bold
			style.Font.IsBold = True

			'Apply style to the cell 
			cells("A1").SetStyle(style)

			'Create a datatable and add three columns to it
			Dim dataTable As New DataTable("Products")
			dataTable.Columns.Add("Product ID", GetType(Int32))
			dataTable.Columns.Add("Product Name", GetType(String))
			dataTable.Columns.Add("Units In Stock", GetType(Int32))

			'Add the first record to it
			Dim dr As DataRow = dataTable.NewRow()
			dr(0) = 1
			dr(1) = "Aniseed Syrup"
			dr(2) = 15
			dataTable.Rows.Add(dr)

			'Add the second record to it
			dr = dataTable.NewRow()
			dr(0) = 2
			dr(1) = "Boston Crab Meat"
			dr(2) = 123
			dataTable.Rows.Add(dr)

			'Import the datatable to sheet cells
			sheet.Cells.ImportDataTable(dataTable, True, "A2")

			'Autofit all the columns in the sheet
			sheet.AutoFitColumns()
		End Sub

		Private Sub ImportDataView(ByVal sheet As Worksheet)
			'Get the cells collection in the worksheet
			Dim cells As Cells = sheet.Cells

			'Put a string value into a cell
			sheet.Cells("A1").PutValue("Import a DataView")

			'Get Style Object 
			Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

			'the font text is set to bold
			style.Font.IsBold = True

			'Apply style to the cell 
			cells("A1").SetStyle(style)

			'Create a datatable and add three columns to it
			Dim dataTable As New DataTable("Products")
			dataTable.Columns.Add("Product ID", GetType(Int32))
			dataTable.Columns.Add("Product Name", GetType(String))
			dataTable.Columns.Add("Units In Stock", GetType(Int32))

			'Add the first record to it
			Dim dr As DataRow = dataTable.NewRow()
			dr(0) = 1
			dr(1) = "Aniseed Syrup"
			dr(2) = 15
			dataTable.Rows.Add(dr)

			'Add the second record to it
			dr = dataTable.NewRow()
			dr(0) = 2
			dr(1) = "Boston Crab Meat"
			dr(2) = 123
			dataTable.Rows.Add(dr)

			'Import the dataview to the sheet cells
			sheet.Cells.ImportDataView(dataTable.DefaultView, True, 1, 0, False)

			'Autofit all the columns in the sheet
			sheet.AutoFitColumns()
		End Sub

		Private Sub ImportFormulaArray(ByVal sheet As Worksheet)
			'Get the cells collection in the worksheet
			Dim cells As Cells = sheet.Cells

			'Put a string value into a cell
			sheet.Cells("A1").PutValue("Import a formula Array")

			'Get Style Object 
			Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

			'the font text is set to bold
			style.Font.IsBold = True

			'Apply style to the cell 
			cells("A1").SetStyle(style)

			'Create a string array and fill it with some formula values
			Dim stringArray() As String = { "=LEN(A1)", "=A2*2", "=SUM(A2:A3)" }

			'Import the array to the sheet cells
			sheet.Cells.ImportFormulaArray(stringArray, 1, 0, True)
		End Sub

		Private Sub ImportFromDataReader(ByVal sheet As Worksheet)
			'Get the cells collection in the worksheet
			Dim cells As Cells = sheet.Cells

			'Put the string value into a cell
			sheet.Cells("A1").PutValue("Import from DataReader")

			'Get Style Object 
			Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

			'the font text is set to bold
			style.Font.IsBold = True

			'Apply style to the cell 
			cells("A1").SetStyle(style)

			Dim path As String = Server.MapPath("~")
			path = path.Substring(0, path.LastIndexOf("\")) & "\Database\Northwind.mdb"
			Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path
			Dim sql As String = "SELECT Country,EmployeeID,FirstName,LastName FROM Employees ORDER BY Country,EmployeeID"

			'Define connection scope
			Using conn As New OleDbConnection(connectionString)
				'Create command object
				Dim command As New OleDbCommand(sql, conn)

				'Open connection
				conn.Open()

				'Create and fill data reader object
				Dim reader As OleDbDataReader
				reader = command.ExecuteReader()

				'Import the datareader object to the sheet cells
				sheet.Cells.ImportDataReader(reader, True, 1, 0, False)

				'sheet.Cells.ImportFromDataReader(reader, true, 1, 0, false);

				' Always call Close when done reading.
				reader.Close()
			End Using

			'Autofit all the columns in the sheet
			sheet.AutoFitColumns()

		End Sub

		Private Sub ImportObjectArray(ByVal sheet As Worksheet)
			'Get the cells collection in the worksheet
			Dim cells As Cells = sheet.Cells

			'Put a string value into a cell
			sheet.Cells("A1").PutValue("Import an object Array")

			'Get Style Object 
			Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

			'the font text is set to bold
			style.Font.IsBold = True

			'Apply style to the cell 
			cells("A1").SetStyle(style)

			'Create an object array and fill it with some values
			Dim obj() As Object = { "Tom", "John", "kelly", 1, 2, 2.8, 5.16, True, False }

			'Import the object array to the sheet cells
			sheet.Cells.ImportObjectArray(obj, 1, 0, False)

			'Autofit all the columns in the sheet
			sheet.AutoFitColumns()
		End Sub

		Private Sub ImportTwoDimensionArray(ByVal sheet As Worksheet)
			'Get the cells collection in the worksheet
			Dim cells As Cells = sheet.Cells

			'Put a string value into a cell
			sheet.Cells("A1").PutValue("Import a two-dimension object Array")

			'Get Style Object 
			Dim style As Aspose.Cells.Style = cells("A1").GetStyle()

			'the font text is set to bold
			style.Font.IsBold = True

			'Apply style to the cell 
			cells("A1").SetStyle(style)

			'Create a multi-dimensional array and fill some values
			Dim objs(1, 2) As Object
			objs(0, 0) = "Product ID"
			objs(0, 1) = 1
			objs(0, 2) = 2
			objs(1, 0) = "Product Name"
			objs(1, 1) = "Aniseed Syrup"
			objs(1, 2) = "Boston Crab Meat"
			'Import the multi-dimensional array to the sheet cells
			sheet.Cells.ImportTwoDimensionArray(objs, 1, 0)
			'Autofit the sheet cells
			sheet.AutoFitColumns()
		End Sub
	End Class
End Namespace


