using System;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for CustomerLabels.
    /// </summary>
    public class CustomerLabels : DbBase
    {
        public CustomerLabels(string path)
            : base(path)
        {

        }

        public Workbook CreateCustomerLabels()
        {
            try
            {
                DBInit();
                //Open the connection object
                this.oleDbConnection1.Open();
                //Specify SQL as command text
                this.oleDbSelectCommand1.CommandText = "SELECT CompanyName, Address, City, Region, PostalCode, Country, CustomerID FROM " +
                    "Customers ORDER BY Country, Region";
                //Fill the datatable
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

            //Get a worksheet
            Worksheet sheet = workbook.Worksheets["Sheet4"];
            //Name the worksheet
            sheet.Name = "Customer Labels";
            //Get the cells collection in the worksheet
            Cells cells = sheet.Cells;
            int row = 0;
            byte column = 0;
            for (int i = 0; i < this.dataTable1.Rows.Count; i++)
            {
                int remainder = i % 3;
                Cell cell;
                switch (remainder)
                {
                    case 0:
                        column = 0;
                        break;
                    case 1:
                        column = 3;
                        break;
                    case 2:
                        column = 6;
                        break;
                }
                //Get a cell
                cell = cells[row, column];
                //Put a value into it
                cell.PutValue((string)this.dataTable1.Rows[i]["CompanyName"]);
                //Get another cell
                cell = cells[row + 1, column];
                //Put a value into it
                cell.PutValue((string)this.dataTable1.Rows[i]["Address"]);
                //Get another cell
                cell = cells[row + 2, column];
                string contact = "";

                if (this.dataTable1.Rows[i]["City"] != DBNull.Value)
                {
                    contact += (string)this.dataTable1.Rows[i]["City"] + " ";
                }
                if (this.dataTable1.Rows[i]["Region"] != DBNull.Value)
                {
                    contact += (string)this.dataTable1.Rows[i]["Region"] + " ";
                }
                if (this.dataTable1.Rows[i]["PostalCode"] != DBNull.Value)
                {
                    contact += (string)this.dataTable1.Rows[i]["PostalCode"];
                }

                //Put the value to it
                cell.PutValue(contact);
                //Get another cell
                cell = cells[row + 3, column];
                //Put a value to it
                cell.PutValue((string)this.dataTable1.Rows[i]["Country"]);

                if (remainder == 2)
                    row += 5;

            }

            //Remove unnecessary worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                sheet = workbook.Worksheets[i];
                if (sheet.Name != "Customer Labels")
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


