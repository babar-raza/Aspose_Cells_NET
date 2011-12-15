using System;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for SalesTotals.
    /// </summary>
    public class SalesTotals : DbBase
    {
        public SalesTotals(string path)
            : base(path)
        {
        }

        public Workbook CreateSalesTotals()
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

            //Open template file
		    string designerFile = MapPath("~/Designer/Northwind.xls");						
            Workbook workbook = new Workbook(designerFile);

            try
            {
                //Specify SQL and execute the query to fill the datatable
                this.oleDbSelectCommand1.CommandText = @"SELECT [Order Subtotals].Subtotal, [Order Subtotals].OrderID, 
				Customers.CompanyName, Customers.CustomerID FROM Customers 
				INNER JOIN ([Order Subtotals] INNER JOIN Orders ON [Order Subtotals].OrderID = Orders.OrderID) 
				ON Customers.CustomerID = Orders.CustomerID 
				WHERE (Orders.ShippedDate BETWEEN #1/1/1995# AND #12/31/1995#) AND ([Order Subtotals].Subtotal > 2500) 
				ORDER BY [Order Subtotals].Subtotal DESC";
                this.oleDbDataAdapter1.Fill(this.dataTable1);
            }
            catch
            {
            }
            finally
            {
                if (this.oleDbConnection1 != null)
                    this.oleDbConnection1.Close();
            }

            //Get the worksheet
            Worksheet sheet = workbook.Worksheets["Sheet12"];
            //Name the worksheet
            sheet.Name = "Sales Totals";
            //Get the cells
            Cells cells = sheet.Cells;
            //Import the datatable to the sheet
            cells.ImportDataTable(this.dataTable1, false, 3, 1, this.dataTable1.Rows.Count, 3);

            decimal totalSum = 0.0m;
            //Input some value to the cells
            for (int i = 0; i < this.dataTable1.Rows.Count; i++)
            {
                totalSum += (decimal)this.dataTable1.Rows[i]["Subtotal"];
                cells[3 + i, 5].PutValue(i + 1);
            }

            //Input value and create a style to apply it
            cells[3 + this.dataTable1.Rows.Count, 0].PutValue("Total:");
            Style style = workbook.Styles[workbook.Styles.Add()];
            style.Font.IsBold = true;
            cells[3 + this.dataTable1.Rows.Count, 0].SetStyle(style);
            //Input a value
            cells[3 + this.dataTable1.Rows.Count, 1].PutValue((double)totalSum);
            //Remove the unnecessary worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                sheet = workbook.Worksheets[i];
                if (sheet.Name != "Sales Totals")
                {
                    workbook.Worksheets.RemoveAt(i);
                    i--;
                }
            }
            //Get the generated workbook
            return workbook;
        }

    }
}
