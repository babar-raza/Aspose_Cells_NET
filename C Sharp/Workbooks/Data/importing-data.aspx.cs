using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using Aspose.Cells;
using System.Data.OleDb;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for ImportingData.
    /// </summary>
    public class ImportingData : System.Web.UI.Page
    {
        protected System.Web.UI.WebControls.DropDownList ImportingDataType;
        protected System.Web.UI.WebControls.Button btnCreateReport;
        protected System.Web.UI.WebControls.DropDownList ddlFileVersion;

        private void Page_Load(object sender, System.EventArgs e)
        {
            // Put user code to initialize the page here
        }

        #region Web Form Designer generated code
        override protected void OnInit(EventArgs e)
        {
            //
            // CODEGEN: This call is required by the ASP.NET Web Form Designer.
            //
            InitializeComponent();
            base.OnInit(e);
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnCreateReport.Click += new System.EventHandler(this.btnCreateReport_Click);
            this.Load += new System.EventHandler(this.Page_Load);

        }
        #endregion

        private void btnCreateReport_Click(object sender, System.EventArgs e)
        {
            //Instantiate a new workbook
            Workbook workbook = new Workbook();
            
            //Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];
            
            //switch case for dropdown's selected value
            switch (ImportingDataType.SelectedItem.Text)
            {
                //if selected text is "Array"
                case "Array":
                    ImportArray(sheet);
                    break;
                //if selected text is "ArrayList"
                case "ArrayList":
                    ImportArrayList(sheet);
                    break;
                //if selected text is "DataColumn"
                case "DataColumn":
                    ImportDataColumn(sheet);
                    break;
                //if selected text is "DataGrid"
                case "DataGrid":
                    ImportDataGrid(sheet);
                    break;
                //if selected text is "DataTable"
                case "DataTable":
                    ImportDataTable(sheet);
                    break;
                //if selected text is "DataView"
                case "DataView":
                    ImportDataView(sheet);
                    break;
                //if selected text is "FormulaArray"
                case "FormulaArray":
                    ImportFormulaArray(sheet);
                    break;
                //if selected text is "FromDataReader"
                case "FromDataReader":
                    ImportFromDataReader(sheet);
                    break;
                //if selected text is "ObjectArray"
                case "ObjectArray":
                    ImportObjectArray(sheet);
                    break;
                //if selected text is "TwoDimensionArray"
                case "TwoDimensionArray":
                    ImportTwoDimensionArray(sheet);
                    break;
                default:
                    return;
            }

            if (ddlFileVersion.SelectedItem.Value == "XLS")
            {
                ////Save file and send to client browser using selected format
                workbook.Save(HttpContext.Current.Response, "ImportingData.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
            }
            else
            {
                workbook.Save(HttpContext.Current.Response, "ImportingData.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
            }

            //end response to avoid unneeded html
            HttpContext.Current.Response.End();      

        }

        private void ImportArray(Worksheet sheet)
        {
            //Get the cells collection in the worksheet
            Cells cells = sheet.Cells;
            
            //Put a string value into a cell
            cells["A1"].PutValue("Import an Array");
            
            //Get Style Object 
            Aspose.Cells.Style style = cells["A1"].GetStyle();
            
            //the font text is set to bold
            style.Font.IsBold = true;
            
            //Apply style to the cell 
            cells["A1"].SetStyle(style);
            
            //Create a string array of values
            string[] names = new string[] { "Tom", "John", "Kelly" };
            
            //Import the array to the sheet cells
            sheet.Cells.ImportArray(names, 1, 0, true);
        }

        private void ImportArrayList(Worksheet sheet)
        {
            //Get the cells collection in the worksheet
            Cells cells = sheet.Cells;
            
            //Put a string value into a cell
            cells["A1"].PutValue("Import an ArrayList");
            
            //Get Style Object 
            Aspose.Cells.Style style = cells["A1"].GetStyle();
            
            //the font text is set to bold
            style.Font.IsBold = true;
            
            //Apply style to the cell 
            cells["A1"].SetStyle(style);
            
            //Create an arraylist and fill some values to it
            ArrayList list = new ArrayList();
            list.Add("Tom");
            list.Add("John");
            list.Add("Kelly");
            
            //Import the arraylist to the sheet cells
            sheet.Cells.ImportArrayList(list, 1, 0, true);
        }

        private void ImportDataColumn(Worksheet sheet)
        {
            //Get the cells collection in the worksheet
            Cells cells = sheet.Cells;
            
            //Put a string value to a cell
            cells["A1"].PutValue("Import a DataColumn");
            
            //Get Style Object 
            Aspose.Cells.Style style = cells["A1"].GetStyle();
            
            //the font text is set to bold
            style.Font.IsBold = true;
            
            //Apply style to the cell 
            cells["A1"].SetStyle(style);
            
            //Create a datatable and add three columns to it
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("Product ID", typeof(Int32));
            dataTable.Columns.Add("Product Name", typeof(string));
            dataTable.Columns.Add("Units In Stock", typeof(Int32));

            //Add the first record to it
            DataRow dr = dataTable.NewRow();
            dr[0] = 1;
            dr[1] = "Aniseed Syrup";
            dr[2] = 15;
            dataTable.Rows.Add(dr);

            //Add a second record to it
            dr = dataTable.NewRow();
            dr[0] = 2;
            dr[1] = "Boston Crab Meat";
            dr[2] = 123;
            dataTable.Rows.Add(dr);

            //Import the datacolumn in the datatable to the sheet cells
            sheet.Cells.ImportDataColumn(dataTable, true, 1, 0, 1, false);
        }

        private void ImportDataGrid(Worksheet sheet)
        {
            //Get the cells collection in the worksheet
            Cells cells = sheet.Cells;
            
            //Put a string value into a cell
            sheet.Cells["A1"].PutValue("Import a DataGrid");
            
            //Get Style Object 
            Aspose.Cells.Style style = cells["A1"].GetStyle();
            
            //the font text is set to bold
            style.Font.IsBold = true;
            
            //Apply style to the cell 
            cells["A1"].SetStyle(style);

            //Create a datatable and add three columns to it
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("Product ID", typeof(Int32));
            dataTable.Columns.Add("Product Name", typeof(string));
            dataTable.Columns.Add("Units In Stock", typeof(Int32));

            //Add the first record to it
            DataRow dr = dataTable.NewRow();
            dr[0] = 1;
            dr[1] = "Aniseed Syrup";
            dr[2] = 15;
            dataTable.Rows.Add(dr);

            //Add the second record to it
            dr = dataTable.NewRow();
            dr[0] = 2;
            dr[1] = "Boston Crab Meat";
            dr[2] = 123;
            dataTable.Rows.Add(dr);

            //Create a datagrid
            DataGrid dataGrid = new DataGrid();
            
            //set its datasource
            dataGrid.DataSource = dataTable;
            
            //bind data
            dataGrid.DataBind();

            //Import the datagrid to sheet cells
            sheet.Cells.ImportDataGrid(dataGrid, 1, 0, false);
            
            //Autofit all the columns in the sheet
            sheet.AutoFitColumns();
        }

        private void ImportDataTable(Worksheet sheet)
        {
            //Get the cells collection in the worksheet
            Cells cells = sheet.Cells;
            
            //Put a string value into a cell
            sheet.Cells["A1"].PutValue("Import a DataTable");
            
            //Get Style Object 
            Aspose.Cells.Style style = cells["A1"].GetStyle();
            
            //the font text is set to bold
            style.Font.IsBold = true;
            
            //Apply style to the cell 
            cells["A1"].SetStyle(style);

            //Create a datatable and add three columns to it
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("Product ID", typeof(Int32));
            dataTable.Columns.Add("Product Name", typeof(string));
            dataTable.Columns.Add("Units In Stock", typeof(Int32));

            //Add the first record to it
            DataRow dr = dataTable.NewRow();
            dr[0] = 1;
            dr[1] = "Aniseed Syrup";
            dr[2] = 15;
            dataTable.Rows.Add(dr);

            //Add the second record to it
            dr = dataTable.NewRow();
            dr[0] = 2;
            dr[1] = "Boston Crab Meat";
            dr[2] = 123;
            dataTable.Rows.Add(dr);

            //Import the datatable to sheet cells
            sheet.Cells.ImportDataTable(dataTable, true, "A2");
            
            //Autofit all the columns in the sheet
            sheet.AutoFitColumns();
        }

        private void ImportDataView(Worksheet sheet)
        {
            //Get the cells collection in the worksheet
            Cells cells = sheet.Cells;
            
            //Put a string value into a cell
            sheet.Cells["A1"].PutValue("Import a DataView");
            
            //Get Style Object 
            Aspose.Cells.Style style = cells["A1"].GetStyle();
           
            //the font text is set to bold
            style.Font.IsBold = true;
            
            //Apply style to the cell 
            cells["A1"].SetStyle(style);

            //Create a datatable and add three columns to it
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("Product ID", typeof(Int32));
            dataTable.Columns.Add("Product Name", typeof(string));
            dataTable.Columns.Add("Units In Stock", typeof(Int32));

            //Add the first record to it
            DataRow dr = dataTable.NewRow();
            dr[0] = 1;
            dr[1] = "Aniseed Syrup";
            dr[2] = 15;
            dataTable.Rows.Add(dr);

            //Add the second record to it
            dr = dataTable.NewRow();
            dr[0] = 2;
            dr[1] = "Boston Crab Meat";
            dr[2] = 123;
            dataTable.Rows.Add(dr);

            //Import the dataview to the sheet cells
            sheet.Cells.ImportDataView(dataTable.DefaultView, true, 1, 0, false);
            
            //Autofit all the columns in the sheet
            sheet.AutoFitColumns();
        }

        private void ImportFormulaArray(Worksheet sheet)
        {
            //Get the cells collection in the worksheet
            Cells cells = sheet.Cells;
            
            //Put a string value into a cell
            sheet.Cells["A1"].PutValue("Import a formula Array");
            
            //Get Style Object 
            Aspose.Cells.Style style = cells["A1"].GetStyle();
            
            //the font text is set to bold
            style.Font.IsBold = true;
            
            //Apply style to the cell 
            cells["A1"].SetStyle(style);

            //Create a string array and fill it with some formula values
            string[] stringArray = { "=LEN(A1)", "=A2*2", "=SUM(A2:A3)" };
            
            //Import the array to the sheet cells
            sheet.Cells.ImportFormulaArray(stringArray, 1, 0, true);
        }

        private void ImportFromDataReader(Worksheet sheet)
        {
            //Get the cells collection in the worksheet
            Cells cells = sheet.Cells;
            
            //Put the string value into a cell
            sheet.Cells["A1"].PutValue("Import from DataReader");
            
            //Get Style Object 
            Aspose.Cells.Style style = cells["A1"].GetStyle();
            
            //the font text is set to bold
            style.Font.IsBold = true;
            
            //Apply style to the cell 
            cells["A1"].SetStyle(style);

            string path = Server.MapPath("~");            
            path = path.Substring(0, path.LastIndexOf("\\")) + "\\Database\\Northwind.mdb";
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path;
            string sql = "SELECT Country,EmployeeID,FirstName,LastName FROM Employees ORDER BY Country,EmployeeID";

            //Define connection scope
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                //Create command object
                OleDbCommand command = new OleDbCommand(sql, conn);
                
                //Open connection
                conn.Open();

                //Create and fill data reader object
                OleDbDataReader reader;
                reader = command.ExecuteReader();
                
                //Import the datareader object to the sheet cells
                sheet.Cells.ImportDataReader(reader, true, 1, 0, false);
                
                //sheet.Cells.ImportFromDataReader(reader, true, 1, 0, false);

                // Always call Close when done reading.
                reader.Close();
            }

            //Autofit all the columns in the sheet
            sheet.AutoFitColumns();

        }

        private void ImportObjectArray(Worksheet sheet)
        {
            //Get the cells collection in the worksheet
            Cells cells = sheet.Cells;
            
            //Put a string value into a cell
            sheet.Cells["A1"].PutValue("Import an object Array");
            
            //Get Style Object 
            Aspose.Cells.Style style = cells["A1"].GetStyle();
            
            //the font text is set to bold
            style.Font.IsBold = true;
            
            //Apply style to the cell 
            cells["A1"].SetStyle(style);
            
            //Create an object array and fill it with some values
            object[] obj = { "Tom", "John", "kelly", 1, 2, 2.8, 5.16, true, false };
            
            //Import the object array to the sheet cells
            sheet.Cells.ImportObjectArray(obj, 1, 0, false);
            
            //Autofit all the columns in the sheet
            sheet.AutoFitColumns();
        }

        private void ImportTwoDimensionArray(Worksheet sheet)
        {
            //Get the cells collection in the worksheet
            Cells cells = sheet.Cells;
            
            //Put a string value into a cell
            sheet.Cells["A1"].PutValue("Import a two-dimension object Array");
            
            //Get Style Object 
            Aspose.Cells.Style style = cells["A1"].GetStyle();
            
            //the font text is set to bold
            style.Font.IsBold = true;
            
            //Apply style to the cell 
            cells["A1"].SetStyle(style);
            
            //Create a multi-dimensional array and fill some values
            object[,] objs = new object[2, 3];
            objs[0, 0] = "Product ID";
            objs[0, 1] = 1;
            objs[0, 2] = 2;
            objs[1, 0] = "Product Name";
            objs[1, 1] = "Aniseed Syrup";
            objs[1, 2] = "Boston Crab Meat";
            //Import the multi-dimensional array to the sheet cells
            sheet.Cells.ImportTwoDimensionArray(objs, 1, 0);
            //Autofit the sheet cells
            sheet.AutoFitColumns();
        }
    }
}


