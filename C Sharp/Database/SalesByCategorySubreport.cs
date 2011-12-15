using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
//using System;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for SalesByCategorySubreport.
    /// </summary>
    public class SalesByCategorySubreport : DbBase
    {
        public SalesByCategorySubreport(string path)
            : base(path)
        {

        }

        public Workbook CreateSalesByCategorySubreport()
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
                //Specify SQL and execute the query to fill a datatable
                this.oleDbSelectCommand1.CommandText = @"SELECT DISTINCTROW Products.ProductName, Sum([Order Details Extended].ExtendedPrice) AS ProductSales
				FROM Categories INNER JOIN (Products INNER JOIN (Orders INNER JOIN [Order Details Extended] ON Orders.OrderID = [Order Details Extended].OrderID) ON Products.ProductID = [Order Details Extended].ProductID) ON Categories.CategoryID = Products.CategoryID
				WHERE (((Orders.OrderDate) Between #1/1/1995# And #12/31/1995#))
				GROUP BY Categories.CategoryID, Categories.CategoryName, Products.ProductName
				ORDER BY Products.ProductName";
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

            //Get a worksheet
            Worksheet sheet = workbook.Worksheets["Sheet9"];
            //Name the sheet
            sheet.Name = "Sales By Category Subreport";
            //Get the cells collection
            Cells cells = sheet.Cells;
            //Import the datatable to the sheet
            cells.ImportDataTable(this.dataTable1, false, 0, 0);
            //Remove the unnecessary worksheets
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                sheet = workbook.Worksheets[i];
                if (sheet.Name != "Sales By Category Subreport")
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


