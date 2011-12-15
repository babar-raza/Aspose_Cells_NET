using System;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for SalesByYearSubreport.
    /// </summary>
    public class SalesByYearSubreport : DbBase
    {
        public SalesByYearSubreport(string path)
            : base(path)
        {
        }

        public Workbook CreateSalesByYearSubreport()
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

            //Open the template file
		    string designerFile = MapPath("~/Designer/Northwind.xls");						
            Workbook workbook = new Workbook(designerFile);

            try
            {
                //Specify an SQL and execute the query to fill the datatable
                this.oleDbSelectCommand1.CommandText = @"SELECT
					DISTINCTROW COUNT(Orders.OrderID) AS Orders, 
					SUM([Order Subtotals].Subtotal) AS Sales, 
					FORMAT(ORDERS.SHIPPEDDATE, 
					'yyyy/Q') AS Quarter
				FROM
					Orders INNER JOIN [Order Subtotals] 
				ON
					Orders.OrderID = [Order Subtotals].OrderID
				WHERE
					(orders.shippeddate IS NOT NULL) GROUP BY FORMAT(ORDERS.SHIPPEDDATE ,  'yyyy/Q')";
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
            Worksheet sheet = workbook.Worksheets["Sheet11"];
            //Set its name
            sheet.Name = "Sales By Year Subreport";
            //Get the cells
            Cells cells = sheet.Cells;

            int currentRow = 0;
            int totalOrders = 0;
            decimal totalSales = 0.0m;
            string thisYear = "";
            SetSalesByYearSubreportStyles(workbook);
            for (int i = 0; i < this.dataTable1.Rows.Count; i++)
            {
                if (i == 0)
                {
                    thisYear = this.dataTable1.Rows[0]["Quarter"].ToString().Substring(0, 4);
                    CreateSalesByYearSubreportHeader(workbook, cells, 0, thisYear);
                    CreateData(cells, 2, 0);
                    totalOrders += (int)this.dataTable1.Rows[0]["Orders"];
                    totalSales += (decimal)this.dataTable1.Rows[0]["Sales"];
                    currentRow = 3;
                }
                else
                {
                    if (thisYear == this.dataTable1.Rows[i]["Quarter"].ToString().Substring(0, 4))
                    {
                        CreateData(cells, currentRow, i);
                        totalOrders += (int)this.dataTable1.Rows[i]["Orders"];
                        totalSales += (decimal)this.dataTable1.Rows[i]["Sales"];
                        currentRow++;
                        if (i == this.dataTable1.Rows.Count - 1)
                        {
                            CreateFooter(workbook, cells, currentRow, totalOrders, totalSales);
                        }
                    }
                    else
                    {
                        CreateFooter(workbook, cells, currentRow, totalOrders, totalSales);
                        totalOrders = 0;
                        totalSales = 0.0m;
                        currentRow++;
                        thisYear = this.dataTable1.Rows[i]["Quarter"].ToString().Substring(0, 4);
                        if (i != this.dataTable1.Rows.Count - 1)
                        {
                            CreateSalesByYearSubreportHeader(workbook, cells, currentRow, thisYear);
                            currentRow += 2;
                            CreateData(cells, currentRow, i);
                            totalOrders += (int)this.dataTable1.Rows[i]["Orders"];
                            totalSales += (decimal)this.dataTable1.Rows[i]["Sales"];
                            currentRow++;
                        }
                    }
                }
            }
            //Remove the unnecessary worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                sheet = workbook.Worksheets[i];
                if (sheet.Name != "Sales By Year Subreport")
                {
                    workbook.Worksheets.RemoveAt(i);
                    i--;
                }
            }
            //Get the generated workbook
            return workbook;
        }

        private void CreateFooter(Workbook workbook, Cells cells, int startRow, int totalOrders, decimal totalSales)
        {
            
            //Get the style
            Style style = workbook.Styles["Bold"];
            //Put value and apply style
            cells[startRow, 1].PutValue("Totals:");
            cells[startRow, 1].SetStyle(style);
            //Put values to cells
            cells[startRow, 2].PutValue(totalOrders);
            cells[startRow, 3].PutValue((double)totalSales);
        }

        private void CreateData(Cells cells, int startRow, int index)
        {
            //Input some values to the cells
            cells[startRow, 1].PutValue(int.Parse(this.dataTable1.Rows[index]["Quarter"].ToString().Substring(5)));
            cells[startRow, 2].PutValue((int)this.dataTable1.Rows[index]["Orders"]);
            cells[startRow, 3].PutValue((double)(decimal)this.dataTable1.Rows[index]["Sales"]);
        }
        private void SetSalesByYearSubreportStyles(Workbook workbook)
        {
            //Create style and specify formatting attributes
            int styleIndex = workbook.Styles.Add();
            Style style = workbook.Styles[styleIndex];
            style.Font.IsBold = true;
            style.Name = "Bold";
        }
        private void CreateSalesByYearSubreportHeader(Workbook workbook, Cells cells, int startRow, string year)
        {
            //Input values and apply styles
            Style style = workbook.Styles["Bold"];
            cells[startRow, 0].PutValue(year + " Summary");
            cells[startRow + 1, 1].PutValue("Quarter:");
            cells[startRow + 1, 2].PutValue("Orders Shipped:");
            cells[startRow + 1, 3].PutValue("Sales:");
            cells[startRow, 0].SetStyle(style);
            cells[startRow + 1, 1].SetStyle(style);
            cells[startRow + 1, 2].SetStyle(style);
            cells[startRow + 1, 3].SetStyle(style);

        }

    }
}


