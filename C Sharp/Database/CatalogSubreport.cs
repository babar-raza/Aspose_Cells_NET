using System;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for CatalogSubreport.
    /// </summary>
    public class CatalogSubreport : DbBase
    {
        public CatalogSubreport(string path)
            : base(path)
        {

        }

        public Workbook CreateCatalogSubreport()
        {
            try
            {
                DBInit();

                //Open the connection object
                this.oleDbConnection1.Open();
                //Specify an SQL as command text
                this.oleDbSelectCommand1.CommandText = "SELECT ProductName, ProductID, QuantityPerUnit, UnitPrice FROM Products ORDER BY " +
                    "ProductName";
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

            //Open a template file
	    string designerFile = MapPath("~/Designer/Northwind.xls");
        Workbook workbook = new Workbook(designerFile);
  
            //Get the sheet
            Worksheet sheet = workbook.Worksheets["Sheet3"];
            //Name the sheet
            sheet.Name = "Catalog Subreport";
            //Get the cells in the sheet
            Cells cells = sheet.Cells;
            //Import the datatable to the sheet
            cells.ImportDataTable(this.dataTable1, false, 0, 1);

            //Remove the unnecessary worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                sheet = workbook.Worksheets[i];
                if (sheet.Name != "Catalog Subreport")
                {
                    workbook.Worksheets.RemoveAt(i);
                    i--;
                }

            }
            //Retrun the generated workbook
            return workbook;
        }
    }
}


