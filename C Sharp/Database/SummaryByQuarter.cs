using System;
using System.Data;
namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for SummaryByQuarter.
    /// </summary>
    public class SummaryByQuarter : DbBase
    {
        public SummaryByQuarter(string path)
            : base(path)
        {

        }

        public Workbook CreateSummaryByQuarter()
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
                //Specify SQL and execute the query to fill the datatable
                this.oleDbSelectCommand1.CommandText = @"SELECT COUNT(Orders.OrderID) AS Orders, 
					SUM([Order Subtotals].Subtotal) AS Sales, FORMAT(Orders.ShippedDate, 'yyyy/Q') AS Quarter 
					FROM Orders INNER JOIN [Order Subtotals] ON Orders.OrderID = [Order Subtotals].OrderID 
					WHERE (Orders.ShippedDate IS NOT NULL) GROUP BY FORMAT(Orders.ShippedDate, 'yyyy/Q')";
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
            Worksheet sheet = workbook.Worksheets["Sheet13"];
            //Name the sheet
            sheet.Name = "Summary By Quarter";
            //Get the cells
            Cells cells = sheet.Cells;

            //Create an arry of datatable with fields
            DataTable[] quarterSummary = new DataTable[4];
            for (int i = 0; i < 4; i++)
            {
                quarterSummary[i] = new DataTable();
                quarterSummary[i].Columns.Add("YearOrQuarter", typeof(int));
                quarterSummary[i].Columns.Add("Orders", typeof(int));
                quarterSummary[i].Columns.Add("Sales", typeof(decimal));
            }

            //Adding some records to the datatables
            for (int i = 0; i < this.dataTable1.Rows.Count; i++)
            {
                string strQuarter = (string)this.dataTable1.Rows[i]["Quarter"];
                int quarter = int.Parse(strQuarter.Substring(strQuarter.Length - 1));
                DataRow row = quarterSummary[quarter - 1].NewRow();
                row["YearOrQuarter"] = int.Parse(strQuarter.Substring(0, 4));
                row["Sales"] = this.dataTable1.Rows[i]["Sales"];
                row["Orders"] = this.dataTable1.Rows[i]["Orders"];
                quarterSummary[quarter - 1].Rows.Add(row);
            }

            //Replace some values in the workbook
            for (int i = 0; i < 4; i++)
            {
                workbook.Replace("&summary" + (i + 1).ToString(), quarterSummary[i]);
            }
            //Remove the unnecessary worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                sheet = workbook.Worksheets[i];
                if (sheet.Name != "Summary By Quarter")
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


