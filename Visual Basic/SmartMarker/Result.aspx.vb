Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Configuration
Imports System.Collections
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Data.OleDb

Namespace Aspose.Cells.Demos.SmartMarker
	Partial Public Class Result
		Inherits System.Web.UI.Page
		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

		End Sub
		Protected Sub btnProcess_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create a dataset based on the custom method
			Dim ds As DataSet = CreateDataSource()

			'Open the template file which contains smart markers
			Dim path As String = MapPath("~/Designer/SmartMarkerDesigner.xls")

			'Create a workbookdesigner object
			Dim designer As New WorkbookDesigner()
			designer.Workbook = New Workbook(path)

			'Set dataset as the datasource
			designer.SetDataSource(ds)
			'Set variable object as another datasource
			designer.SetDataSource("Variable", "Single Variable")
			'Set multi-valued variable array as another datasource
			designer.SetDataSource("MultiVariable", New String() { "Variable 1", "Variable 2", "Variable 3" })
			'Set multi-valued variable array as another datasource
			designer.SetDataSource("MultiVariable2", New String() { "Skip 1", "Skip 2", "Skip 3" })

			'Process the smart markers in the designer file
			designer.Process()

			'Save the excel file
			designer.Workbook.Save(HttpContext.Current.Response,"SmartMarker.xls", ContentDisposition.Attachment,New XlsSaveOptions(SaveFormat.Excel97To2003))
		End Sub

		#Region "Private code to create data source"

		Private Function CreateDataSource() As DataSet
			'Using ADO.NET APIs

			'Create a dataset
			Dim ds As New DataSet()
			'Create a connection object
			Dim oleDbConnection1 As New OleDbConnection()
			Try
				'Set the connection string and specify the database file path
				Dim path As String = MapPath(".")
				path = path.Substring(0, path.LastIndexOf("\"))
				oleDbConnection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & "\Database\Northwind.mdb"

				'Create a dataadapter object with specific set of attributes
				Dim oleDbDataAdapter1 As New OleDbDataAdapter()
				oleDbDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() { New System.Data.Common.DataTableMapping("Table", "Order Details", New System.Data.Common.DataColumnMapping() { New System.Data.Common.DataColumnMapping("Discount", "Discount"), New System.Data.Common.DataColumnMapping("OrderID", "OrderID"), New System.Data.Common.DataColumnMapping("ProductID", "ProductID"), New System.Data.Common.DataColumnMapping("Quantity", "Quantity"), New System.Data.Common.DataColumnMapping("UnitPrice", "UnitPrice") }) })
				'Create a command object
				Dim oleDbSelectCommand1 As New OleDbCommand()
				'Specify the connection object
				oleDbSelectCommand1.Connection = oleDbConnection1
				'Specify the command object for execution
				oleDbDataAdapter1.SelectCommand = oleDbSelectCommand1
				'Specify the SQL command text
				oleDbSelectCommand1.CommandText = "SELECT Discount, OrderID, ProductID, Quantity, UnitPrice FROM [Order Details]"

				'Create another dataadapter object with specific set of attributes
				Dim oleDbDataAdapter2 As New OleDbDataAdapter()
				oleDbDataAdapter2.TableMappings.AddRange(New System.Data.Common.DataTableMapping() { New System.Data.Common.DataTableMapping("Table", "Customers", New System.Data.Common.DataColumnMapping() { New System.Data.Common.DataColumnMapping("Address", "Address"), New System.Data.Common.DataColumnMapping("City", "City"), New System.Data.Common.DataColumnMapping("CompanyName", "CompanyName"), New System.Data.Common.DataColumnMapping("ContactName", "ContactName"), New System.Data.Common.DataColumnMapping("ContactTitle", "ContactTitle"), New System.Data.Common.DataColumnMapping("Country", "Country"), New System.Data.Common.DataColumnMapping("CustomerID", "CustomerID"), New System.Data.Common.DataColumnMapping("Fax", "Fax"), New System.Data.Common.DataColumnMapping("Phone", "Phone"), New System.Data.Common.DataColumnMapping("PostalCode", "PostalCode"), New System.Data.Common.DataColumnMapping("Region", "Region") }) })

				'Create another command object
				Dim oleDbSelectCommand2 As New OleDbCommand()
				'Specify the connection object
				oleDbSelectCommand2.Connection = oleDbConnection1
				'Specify the command object for execution
				oleDbDataAdapter2.SelectCommand = oleDbSelectCommand2
				'Specify the SQL command text
				oleDbSelectCommand2.CommandText = "SELECT Address, City, CompanyName, ContactName, ContactTitle, Country, CustomerID, Fax, Phone, PostalCode, Region FROM Customers"
				'Open the connection
				oleDbConnection1.Open()
				'Fill the dataset based on the dataadapter objects
				oleDbDataAdapter1.Fill(ds)
				oleDbDataAdapter2.Fill(ds)
			Catch
			Finally
				'Close the connection object
				oleDbConnection1.Close()
			End Try

			'Return the dataset object
			Return ds
		End Function
		#End Region

	End Class
End Namespace
