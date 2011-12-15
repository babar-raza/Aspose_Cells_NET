using System;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for SalesByYear.
    /// </summary>
    public class SalesByYear : DbBase
    {
        public SalesByYear(string path)
            : base(path)
        {

        }

        public Workbook CreateSalesByYear()
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

            //Specify an SQL and execute the query to fill a datatable
            this.oleDbSelectCommand1.CommandText = @"SELECT DISTINCTROW Format([ShippedDate],""yyyy-mm-dd"") AS [ShippedDate], Orders.OrderID, [Order Subtotals].Subtotal as Subtotal
				FROM Orders INNER JOIN [Order Subtotals] ON Orders.OrderID = [Order Subtotals].OrderID
				WHERE( (Orders.ShippedDate) Is Not Null)";
            this.oleDbDataAdapter1.Fill(this.dataTable1);

            //Get the sheet
            Worksheet sheet = workbook.Worksheets["Sheet10"];
            //Name the sheet
            sheet.Name = "Sales By Year";
            //Get the cells collection
            Cells cells = sheet.Cells;
            //Import the datatable to the sheet
            cells.ImportDataTable(this.dataTable1, false, 6, 2);
            //Input values to some cells
            for (int i = 0; i < this.dataTable1.Rows.Count; i++)
                cells[6 + i, 1].PutValue(i + 1);
            //Remove the unnecessary worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                sheet = workbook.Worksheets[i];
                if (sheet.Name != "Sales By Year")
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


