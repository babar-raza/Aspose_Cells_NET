using System;
using System.Data;
using System.Drawing;
using System.Data.OleDb;
namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for Catalog.
    /// </summary>
    public class Catalog : DbBase
    {
        public Catalog(string path)
            : base(path)
        {

        }

        public Workbook CreateCatalog()
        {

            try
            {
                DBInit();
            }
            catch
            {
            }

            //Open a template file
	    string designerFile = MapPath("~/Designer/Northwind.xls");
        Workbook workbook = new Workbook(designerFile);
 

            ReadCategory();
            //Create a new datatable
            DataTable dataTable2 = new DataTable();
            //Get a worksheet
            Worksheet sheet = workbook.Worksheets["Sheet2"];
            //Name the sheet
            sheet.Name = "Catalog";
            //Get the worksheet cells
            Cells cells = sheet.Cells;

            int currentRow = 55;

            //Add LightGray color to color palette
            workbook.ChangePalette(Color.LightGray, 55);
            //Get the workbook's styles collection
            StyleCollection styles = workbook.Styles;
            //Set CategoryName style with formatting attributes
            int styleIndex = styles.Add();
            Style styleCategoryName = styles[styleIndex];
            styleCategoryName.Font.Size = 14;
            styleCategoryName.Font.Color = Color.Blue;
            styleCategoryName.Font.IsBold = true;
            styleCategoryName.Font.Name = "Times New Roman";

            //Set Description style with formatting attributes
            styleIndex = styles.Add();
            Style styleDescription = styles[styleIndex];
            styleDescription.Font.Name = "Times New Roman";
            styleDescription.Font.Color = Color.Blue;
            styleDescription.Font.IsItalic = true;

            //Set ProductName style with formatting attributes
            styleIndex = styles.Add();
            Style styleProductName = styles[styleIndex];
            styleProductName.Font.IsBold = true;

            //Set Title style with formatting attributes
            styleIndex = styles.Add();
            Style styleTitle = styles[styleIndex];
            styleTitle.Font.IsBold = true;
            styleTitle.Font.IsItalic = true;
            styleTitle.ForegroundColor = Color.LightGray;

            styleIndex = styles.Add();
            Style styleNumber = styles[styleIndex];
            styleNumber.Font.Name = "Times New Roman";
            styleNumber.Number = 8;

            //Create the styleflag struct
            StyleFlag styleflag = new StyleFlag();
            styleflag.All = true;
            //Get the horizontal page breaks collection
            HorizontalPageBreakCollection hPageBreaks = sheet.HorizontalPageBreaks;

            //Specify SQL for the command
            string cmd = "SELECT ProductName, ProductID, QuantityPerUnit, " +
                "UnitPrice FROM Products";
            for (int i = 0; i < this.dataTable1.Rows.Count; i++)
            {
                currentRow += 2;
                cells.SetRowHeight(currentRow, 20);
                cells[currentRow, 1].SetStyle(styleCategoryName);
                DataRow categoriesRow = this.dataTable1.Rows[i];

                //Write CategoryName
                cells[currentRow, 1].PutValue((string)categoriesRow["CategoryName"]);

                //Write Description
                currentRow++;
                cells[currentRow, 1].PutValue((string)categoriesRow["Description"]);
                cells[currentRow, 1].SetStyle(styleDescription);

                dataTable2.Clear();

                //Execuate command and fill the datatable
                try
                {
                    this.oleDbDataAdapter2 = new OleDbDataAdapter();
                    string cmdText = cmd + " where categoryid = "
                        + categoriesRow["CategoryID"].ToString();
                    this.oleDbDataAdapter2.SelectCommand = new OleDbCommand(cmdText, this.oleDbConnection1);
                    this.oleDbConnection1.Open();
                    oleDbDataAdapter2.Fill(dataTable2);
                }
                catch
                {
                }
                finally
                {
                    oleDbDataAdapter2.Dispose();
                    this.oleDbConnection1.Close();
                }

                currentRow += 2;
                //Import the datatable to the sheet
                cells.ImportDataTable(dataTable2, true, currentRow, 1);
                //Create a range
                Range range = cells.CreateRange(currentRow, 1, 1, 4);
                //Apply style to the range
                range.ApplyStyle(styleTitle, styleflag);
                //Create a range
                range = cells.CreateRange(currentRow + 1, 1, dataTable2.Rows.Count, 1);
                //Apply style to the range
                range.ApplyStyle(styleProductName, styleflag);
                //Create a range
                range = cells.CreateRange(currentRow + 1, 4, dataTable2.Rows.Count, 1);
                //Apply style to the range
                range.ApplyStyle(styleNumber, styleflag);

                currentRow += dataTable2.Rows.Count;
                //Apply horizontal page breaks
                hPageBreaks.Add(currentRow, 0);
            }

            //Remove the unnecessary worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                sheet = workbook.Worksheets[i];
                if (sheet.Name != "Catalog")
                {
                    workbook.Worksheets.RemoveAt(i);
                    i--;
                }

            }
            //Return the generated workbook
            return workbook;
        }

        private void ReadCategory()
        {
            //Execute the command and fill a datatable
            try
            {
                this.oleDbConnection1.Open();
                this.oleDbSelectCommand1.CommandText = "SELECT CategoryID, CategoryName, Description FROM Categories";
                this.oleDbDataAdapter1.Fill(this.dataTable1);
            }
            catch
            {
            }
            finally
            {
                this.oleDbDataAdapter1.Dispose();
                this.oleDbConnection1.Close();
            }

        }

    }
}


