using System;
using System.Data;
using System.Drawing;
using System.Data.OleDb;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for EmployeeSales.
    /// </summary>
    public class EmployeeSales : DbBase
    {
        public EmployeeSales(string path)
            : base(path)
        {

        }

        public Workbook CreateEmployeeSales()
        {
            try
            {
                DBInit();
            }
            catch
            {
            }
            finally
            {
                if (this.oleDbConnection1 != null)
                    this.oleDbConnection1.Close();
            }


            //Open a template file
	    string designerFile = MapPath("~/Designer/Northwind.xls");						
            Workbook workbook = new Workbook(designerFile);

            //Get the worksheet
            Worksheet sheet = workbook.Worksheets["Sheet5"];
            //Name the worksheet
            sheet.Name = "Employee Sales UK";
            //Get the cells
            Cells cells = sheet.Cells;
            //Get the worksheet
            sheet = workbook.Worksheets["Sheet6"];
            //Get its cells
            cells = sheet.Cells;
            //Name the sheet
            sheet.Name = "Employee Sales USA";

            ReadEmployees();
            //Create datatable array
            DataTable[] dtSales = this.CreateDataResult();

            int currentUKRow = 6;
            int currentUSARow = 6;

            int styleIndex;
            Style style;
            //Create a header style with specific formatting
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
            style.Borders.SetColor(Color.Black);
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.IsTextWrapped = true;
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.Name = "HeaderStyle";

            //Create different styles with specific formattings and apply to different cells
            //Input different values and set formulas to some cells, importing datatables to the cells  
            for (int i = 0; i < this.dataTable1.Rows.Count; i++)
            {
                string employeeName = (string)this.dataTable1.Rows[i]["LastName"] + ","
                    + (string)this.dataTable1.Rows[i]["FirstName"];
                if (this.dataTable1.Rows[i]["Country"].ToString() == "UK")
                {
                    sheet = workbook.Worksheets["Employee Sales UK"];
                    cells = sheet.Cells;

                    cells[currentUKRow - 2, 0].PutValue("Salesperson:" + employeeName);
                    style = workbook.Styles[workbook.Styles.Add()];
                    style.Font.IsBold = true;
                    style.Font.Size = 12;

                    cells[currentUKRow - 2, 0].SetStyle(style);

                    if ((decimal)this.dataTable1.Rows[i]["TotalSales"] > 5000)
                    {
                        cells[currentUKRow - 2, 3].PutValue("Exceeded Goal!");
                        style = workbook.Styles[workbook.Styles.Add()];

                        Font font = style.Font;
                        font.Color = Color.Red;
                        font.IsItalic = true;
                        font.Size = 12;
                        font.IsBold = true;

                        cells[currentUKRow - 2, 3].SetStyle(style);
                    }
                    cells.SetRowHeight(currentUKRow - 2, 19);
                    cells.SetRowHeight(currentUKRow - 1, 4);
                    cells.SetRowHeight(currentUKRow, 48);

                    style = workbook.Styles["HeaderStyle"];
                    for (int j = 1; j < 5; j++)
                        cells[currentUKRow, (byte)j].SetStyle(style);
                    cells[currentUKRow, 1].PutValue("Order ID:");
                    cells[currentUKRow, 2].PutValue("Sales Amount:");
                    cells[currentUKRow, 3].PutValue("Percent of Salesperson's Total:");
                    cells[currentUKRow, 4].PutValue("Percent of Country Total:");
                    currentUKRow++;

                    cells.ImportDataTable(dtSales[i], false, currentUKRow, 1);
                    string startCellName1 = CellsHelper.CellIndexToName(currentUKRow, 2);
                    string startCellName2 = CellsHelper.CellIndexToName(currentUKRow, 4);

                    currentUKRow += dtSales[i].Rows.Count - 1;
                    string endCellName1 = CellsHelper.CellIndexToName(currentUKRow, 2);
                    string endCellName2 = CellsHelper.CellIndexToName(currentUKRow, 4);

                    cells[currentUKRow + 1, 2].Formula = "=sum(" + startCellName1 + ":"
                        + endCellName1 + ")";

                    cells[currentUKRow + 1, 4].Formula = "=sum(" + startCellName2 + ":"
                        + endCellName2 + ")";

                    currentUKRow += 4;

                }
                else
                {
                    sheet = workbook.Worksheets["Employee Sales USA"];
                    cells = sheet.Cells;

                    cells[currentUSARow - 2, 0].PutValue("Salesperson:" + employeeName);

                    style = workbook.Styles[workbook.Styles.Add()];
                    style.Font.IsBold = true;
                    style.Font.Size = 12;
                    cells[currentUSARow - 2, 0].SetStyle(style);

                    if ((decimal)this.dataTable1.Rows[i]["TotalSales"] > 5000)
                    {
                        style = workbook.Styles[workbook.Styles.Add()];

                        cells[currentUSARow - 2, 3].PutValue("Exceeded Goal!");
                        Font font = style.Font;
                        font.Color = Color.Red;
                        font.IsItalic = true;
                        font.Size = 12;
                        font.IsBold = true;

                        cells[currentUSARow - 2, 3].SetStyle(style);
                    }
                    cells.SetRowHeight(currentUSARow - 2, 19);
                    cells.SetRowHeight(currentUSARow - 1, 4);
                    cells.SetRowHeight(currentUSARow, 48);

                    style = workbook.Styles["HeaderStyle"];
                    for (int j = 1; j < 5; j++)
                        cells[currentUSARow, (byte)j].SetStyle(style);
                    cells[currentUSARow, 1].PutValue("Order ID:");
                    cells[currentUSARow, 2].PutValue("Sales Amount:");
                    cells[currentUSARow, 3].PutValue("Percent of Salesperson's Total:");
                    cells[currentUSARow, 4].PutValue("Percent of Country Total:");
                    currentUSARow++;

                    cells.ImportDataTable(dtSales[i], false, currentUSARow, 1);

                    string startCellName1 = CellsHelper.CellIndexToName(currentUSARow, 2);
                    string startCellName2 = CellsHelper.CellIndexToName(currentUSARow, 4);

                    currentUSARow += dtSales[i].Rows.Count - 1;
                    string endCellName1 = CellsHelper.CellIndexToName(currentUSARow, 2);
                    string endCellName2 = CellsHelper.CellIndexToName(currentUSARow, 4);

                    cells[currentUSARow + 1, 2].Formula = "=sum(" + startCellName1 + ":"
                        + endCellName1 + ")";

                    cells[currentUSARow + 1, 4].Formula = "=sum(" + startCellName2 + ":"
                        + endCellName2 + ")";

                    currentUSARow += 4;
                }
            }
            //Remove unnecessary worksheets
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                sheet = workbook.Worksheets[i];
                if (sheet.Name != "Employee Sales UK" && sheet.Name != "Employee Sales USA")
                {
                    workbook.Worksheets.RemoveAt(i);
                    i--;
                }

            }
            //Get the generated workbook
            return workbook;
        }

        private void ReadEmployees()
        {
            try
            {
                //Open the connection
                this.oleDbConnection1.Open();
                //Specify SQL
                this.oleDbSelectCommand1.CommandText = "SELECT Country, EmployeeID, FirstName, LastName FROM Employees ORDER BY Country, " +
                    "LastName, FirstName";
                //Fill a datatable executing the query
                this.oleDbDataAdapter1.Fill(this.dataTable1);
            }
            catch
            {
            }
            finally
            {
                if (this.oleDbDataAdapter1 != null)
                    this.oleDbDataAdapter1.Dispose();
                if (this.oleDbConnection1 != null)
                    this.oleDbConnection1.Close();
            }
            //Add a column to the datatable
            this.dataTable1.Columns.Add("TotalSales", typeof(decimal));

        }

        private DataTable[] CreateDataResult()
        {
            //Create datatable array
            DataTable[] dtSales = new DataTable[this.dataTable1.Rows.Count];
            decimal totalUKSales = 0.0M;
            decimal totalUSASales = 0.0M;
            //Specify SQL
            string cmd = "SELECT Orders.OrderID, [Order Subtotals].Subtotal as SaleAmount FROM Employees INNER " +
                "JOIN (Orders INNER JOIN [Order Subtotals] ON Orders.OrderID = [Order Subtotals].OrderID) " +
                "ON Employees.EmployeeID = Orders.EmployeeID";

            //Create different datatables and fill them with different sets of data based on specific SQL
            for (int i = 0; i < this.dataTable1.Rows.Count; i++)
            {
                dtSales[i] = new DataTable();
                dtSales[i].Columns.Add("OrderID", typeof(int));
                dtSales[i].Columns.Add("SaleAmount", typeof(decimal));
                dtSales[i].Columns.Add("PercentOfPerson", typeof(decimal));
                dtSales[i].Columns.Add("PercentOfCountry", typeof(decimal));

                decimal totalPersonSales = 0.0M;
                try
                {
                    this.oleDbDataAdapter2 = new OleDbDataAdapter();
                    string cmdText = cmd + " where Employees.EmployeeID ="
                        + this.dataTable1.Rows[i]["EmployeeID"].ToString();
                    this.oleDbDataAdapter2.SelectCommand = new OleDbCommand(cmdText, this.oleDbConnection1);
                    this.oleDbConnection1.Open();
                    this.oleDbDataAdapter2.Fill(dtSales[i]);
                }
                catch
                {
                }
                finally
                {
                    if (this.oleDbDataAdapter2 != null)
                        this.oleDbDataAdapter2.Dispose();
                    if (this.oleDbConnection1 != null)
                        this.oleDbConnection1.Close();
                }

                //Get total sales amount of a salesperson
                for (int j = 0; j < dtSales[i].Rows.Count; j++)
                {
                    totalPersonSales += (decimal)dtSales[i].Rows[j]["SaleAmount"];
                }

                this.dataTable1.Rows[i]["TotalSales"] = totalPersonSales;

                //Get the percent
                for (int j = 0; j < dtSales[i].Rows.Count; j++)
                {
                    dtSales[i].Rows[j]["PercentOfPerson"] = (decimal)dtSales[i].Rows[j]["SaleAmount"] / totalPersonSales;
                }
                if (this.dataTable1.Rows[i]["Country"].ToString() == "UK")
                    totalUKSales += totalPersonSales;
                else
                    totalUSASales += totalPersonSales;
            }
            for (int i = 0; i < dtSales.Length; i++)
            {
                for (int j = 0; j < dtSales[i].Rows.Count; j++)
                {
                    if (this.dataTable1.Rows[i]["Country"].ToString() == "UK")
                    {
                        dtSales[i].Rows[j]["PercentOfCountry"] = (decimal)dtSales[i].Rows[j]["SaleAmount"] / totalUKSales;
                    }
                    else
                    {
                        dtSales[i].Rows[j]["PercentOfCountry"] = (decimal)dtSales[i].Rows[j]["SaleAmount"] / totalUSASales;
                    }
                }
            }
            //Get the generated datatables
            return dtSales;
        }

    }
}


