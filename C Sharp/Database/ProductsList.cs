using System;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for ProductsList.
    /// </summary>
    public class ProductsList : DbBase
    {


        public ProductsList(string path)
            : base(path)
        {
        }


        public Workbook CreateProductsList()
        {
            try
            {
                DBInit();
                //Open the connection
                this.oleDbConnection1.Open();
                //Specify an SQL query as command text
                this.oleDbSelectCommand1.CommandText = @"SELECT	DISTINCTROW Products.ProductName, 
														Categories.CategoryName, 
														Products.QuantityPerUnit, 
														Products.UnitsInStock
													FROM Categories INNER JOIN Products 
													ON	Categories.CategoryID = Products.CategoryID
													WHERE
														(((Products.Discontinued) = No))
													Order by Products.ProductName";
                //Fill a datatable
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

            //Open a template excel file
	    string designerFile = MapPath("~/Designer/Northwind.xls");						
            Workbook workbook = new Workbook(designerFile);

            //Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];
            //Import a datatable to the sheet
            sheet.Cells.ImportDataTable(this.dataTable1, false, 6, 1);
            //Name the sheet
            sheet.Name = "Products List";

            //Remove all other worksheets (except the first worksheet) in the workbook
            while (workbook.Worksheets.Count > 1)
                
                workbook.Worksheets.RemoveAt(workbook.Worksheets.Count - 1);

            //Return the generated workbook
            return workbook;
        }


    }
}


