using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.OleDb;

namespace Aspose.Cells.Demos.SmartMarker
{
    public partial class Result : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void btnProcess_Click(object sender, EventArgs e)
        {
            //Create a dataset based on the custom method
            DataSet ds = CreateDataSource();

            //Open the template file which contains smart markers
            string path = MapPath("~/Designer/SmartMarkerDesigner.xls");

            //Create a workbookdesigner object
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook(path);

            //Set dataset as the datasource
            designer.SetDataSource(ds);
            //Set variable object as another datasource
            designer.SetDataSource("Variable", "Single Variable");
            //Set multi-valued variable array as another datasource
            designer.SetDataSource("MultiVariable", new string[] { "Variable 1", "Variable 2", "Variable 3" });
            //Set multi-valued variable array as another datasource
            designer.SetDataSource("MultiVariable2", new string[] { "Skip 1", "Skip 2", "Skip 3" });

            //Process the smart markers in the designer file
            designer.Process();

            //Save the excel file
            designer.Workbook.Save(HttpContext.Current.Response,"SmartMarker.xls", ContentDisposition.Attachment,new XlsSaveOptions(SaveFormat.Excel97To2003));
        }

        #region Private code to create data source

        private DataSet CreateDataSource()
        {
            //Using ADO.NET APIs

            //Create a dataset
            DataSet ds = new DataSet();
            //Create a connection object
            OleDbConnection oleDbConnection1 = new OleDbConnection();
            try
            {
                //Set the connection string and specify the database file path
                string path = MapPath(".");
                path = path.Substring(0, path.LastIndexOf("\\"));
                oleDbConnection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "\\Database\\Northwind.mdb";

                //Create a dataadapter object with specific set of attributes
                OleDbDataAdapter oleDbDataAdapter1 = new OleDbDataAdapter();
                oleDbDataAdapter1.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] { new System.Data.Common.DataTableMapping("Table", "Order Details", new System.Data.Common.DataColumnMapping[] { new System.Data.Common.DataColumnMapping("Discount", "Discount"), new System.Data.Common.DataColumnMapping("OrderID", "OrderID"), new System.Data.Common.DataColumnMapping("ProductID", "ProductID"), new System.Data.Common.DataColumnMapping("Quantity", "Quantity"), new System.Data.Common.DataColumnMapping("UnitPrice", "UnitPrice") }) });
                //Create a command object
                OleDbCommand oleDbSelectCommand1 = new OleDbCommand();
                //Specify the connection object
                oleDbSelectCommand1.Connection = oleDbConnection1;
                //Specify the command object for execution
                oleDbDataAdapter1.SelectCommand = oleDbSelectCommand1;
                //Specify the SQL command text
                oleDbSelectCommand1.CommandText = "SELECT Discount, OrderID, ProductID, Quantity, UnitPrice FROM [Order Details]";

                //Create another dataadapter object with specific set of attributes
                OleDbDataAdapter oleDbDataAdapter2 = new OleDbDataAdapter();
                oleDbDataAdapter2.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] { new System.Data.Common.DataTableMapping("Table", "Customers", new System.Data.Common.DataColumnMapping[] { new System.Data.Common.DataColumnMapping("Address", "Address"), new System.Data.Common.DataColumnMapping("City", "City"), new System.Data.Common.DataColumnMapping("CompanyName", "CompanyName"), new System.Data.Common.DataColumnMapping("ContactName", "ContactName"), new System.Data.Common.DataColumnMapping("ContactTitle", "ContactTitle"), new System.Data.Common.DataColumnMapping("Country", "Country"), new System.Data.Common.DataColumnMapping("CustomerID", "CustomerID"), new System.Data.Common.DataColumnMapping("Fax", "Fax"), new System.Data.Common.DataColumnMapping("Phone", "Phone"), new System.Data.Common.DataColumnMapping("PostalCode", "PostalCode"), new System.Data.Common.DataColumnMapping("Region", "Region") }) });

                //Create another command object
                OleDbCommand oleDbSelectCommand2 = new OleDbCommand();
                //Specify the connection object
                oleDbSelectCommand2.Connection = oleDbConnection1;
                //Specify the command object for execution
                oleDbDataAdapter2.SelectCommand = oleDbSelectCommand2;
                //Specify the SQL command text
                oleDbSelectCommand2.CommandText = "SELECT Address, City, CompanyName, ContactName, ContactTitle, Country, CustomerID, Fax, Phone, PostalCode, Region FROM Customers";
                //Open the connection
                oleDbConnection1.Open();
                //Fill the dataset based on the dataadapter objects
                oleDbDataAdapter1.Fill(ds);
                oleDbDataAdapter2.Fill(ds);
            }
            catch
            {
            }
            finally
            {
                //Close the connection object
                oleDbConnection1.Close();
            }

            //Return the dataset object
            return ds;
        }
        #endregion

    }
}
