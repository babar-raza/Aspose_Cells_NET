using System;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for ProductsByCategory.
    /// </summary>
    public class ProductsByCategory : DbBase
    {
        public ProductsByCategory(string path)
            : base(path)
        {

        }

        public Workbook CreateProductsByCategory()
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


            this.dataTable1.Reset();
            try
            {
                //Specify an SQL and execute the query to fill the datatable
                this.oleDbDataAdapter1.SelectCommand.CommandText = @"SELECT Categories.CategoryName, Products.ProductName, Products.QuantityPerUnit, Products.UnitsInStock, Products.Discontinued, Categories.CategoryID, Products.ProductID FROM Categories INNER JOIN Products ON Categories.CategoryID = Products.CategoryID WHERE (Products.Discontinued <> Yes) ORDER BY Categories.CategoryName, Products.ProductName";
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
            Worksheet sheet = workbook.Worksheets["Sheet7"];
            //Name it
            sheet.Name = "Products By Category";
            //Get the cells
            Cells cells = sheet.Cells;
            //Get the sheet vertical page breaks
            VerticalPageBreakCollection vPageBreaks = sheet.VerticalPageBreaks;
            //Set row heights
            cells.SetRowHeight(4, 20.25);
            cells.SetRowHeight(5, 18.75);
            ushort currentRow = 4;
            byte currentColumn = 0;

            string lastCategory = "";
            string thisCategory, nextCategory;

            int productsCount = 0;

            SetProductsByCategoryStyles(workbook);
            //Fill cells by inputing the values and apply styles to the data
            for (int i = 0; i < this.dataTable1.Rows.Count; i++)
            {
                thisCategory = (string)this.dataTable1.Rows[i]["CategoryName"];
                if (thisCategory != lastCategory)
                {
                    currentRow = 4;
                    if (i != 0)
                        currentColumn += 4;
                    CreateProductsByCategoryHeader(workbook, cells, currentRow, currentColumn, thisCategory);
                    lastCategory = thisCategory;
                    currentRow += 2;
                }
                cells[currentRow, currentColumn].PutValue((string)this.dataTable1.Rows[i]["ProductName"]);
                cells[currentRow, (byte)(currentColumn + 1)].PutValue((short)this.dataTable1.Rows[i]["UnitsInStock"]);

                if (i != this.dataTable1.Rows.Count - 1)
                {
                    nextCategory = (string)this.dataTable1.Rows[i + 1]["CategoryName"];
                    if (thisCategory != nextCategory)
                    {
                        Style style = workbook.Styles["ProductsCount"];
                        cells[currentRow + 1, currentColumn].PutValue("Number of Products:");
                        cells[currentRow + 1, currentColumn].SetStyle(style);

                        style = workbook.Styles["CountNumber"];
                        cells[currentRow + 1, (byte)(currentColumn + 1)].PutValue(productsCount + 1);
                        cells[currentRow + 1, (byte)(currentColumn + 1)].SetStyle(style);
                        currentRow++;
                        productsCount = 0;
                        vPageBreaks.Add(0, currentColumn + 1);
                    }
                    else
                        productsCount++;
                }
                else
                {
                    Style style = workbook.Styles["ProductsCount"];
                    cells[currentRow + 1, currentColumn].PutValue("Number of Products:");
                    cells[currentRow + 1, currentColumn].SetStyle(style);

                    style = workbook.Styles["CountNumber"];
                    cells[currentRow + 1, (byte)(currentColumn + 1)].PutValue(productsCount + 1);
                    cells[currentRow + 1, (byte)(currentColumn + 1)].SetStyle(style);
                }
                currentRow++;
            }

            //Remove the unnecessary worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                sheet = workbook.Worksheets[i];
                if (sheet.Name != "Products By Category")
                {
                    workbook.Worksheets.RemoveAt(i);
                    i--;
                }
            }
            //Get the generated workbook
            return workbook;
        }

        private void SetProductsByCategoryStyles(Workbook workbook)
        {
            //Create a style with some specific formatting attributes
            int styleIndex = workbook.Styles.Add();
            Style style = workbook.Styles[styleIndex];
            style.Font.IsItalic = true;
            style.Font.IsBold = true;
            style.Font.Size = 16;
            style.HorizontalAlignment = TextAlignmentType.Right;
            style.Name = "Category";

            //Create a style with some specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Font.Size = 16;
            style.Font.IsBold = true;
            style.HorizontalAlignment = TextAlignmentType.Left;
            style.Name = "CategoryName";

            //Create a style with some specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Font.Size = 14;
            style.Font.IsBold = true;
            style.Font.IsItalic = true;
            style.HorizontalAlignment = TextAlignmentType.Left;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Medium;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Medium;
            style.Name = "ProductName";

            //Create a style with some specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Font.Size = 14;
            style.Font.IsBold = true;
            style.Font.IsItalic = true;
            style.HorizontalAlignment = TextAlignmentType.Right;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Medium;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Medium;
            style.Name = "UnitsInStock";

            //Create a style with some specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Font.IsBold = true;
            style.Font.IsItalic = true;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Name = "ProductsCount";

            //Create a style with some specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.HorizontalAlignment = TextAlignmentType.Left;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Name = "CountNumber";

        }
        private void CreateProductsByCategoryHeader(Workbook workbook, Cells cells, ushort startRow, byte startColumn, string categoryName)
        {
            //Input values and apply the styles to the cells

            Style style = workbook.Styles["Category"];
            cells[startRow, startColumn].PutValue("Category:");
            cells[startRow, startColumn].SetStyle(style);

            style = workbook.Styles["CategoryName"];
            cells[startRow, (byte)(startColumn + 1)].PutValue(categoryName);
            cells[startRow, (byte)(startColumn + 1)].SetStyle(style);

            style = workbook.Styles["ProductName"];
            cells[startRow + 1, startColumn].PutValue("Product Name");
            cells[startRow + 1, startColumn].SetStyle(style);

            style = workbook.Styles["UnitsInStock"];
            cells[startRow + 1, (byte)(startColumn + 1)].PutValue("Units In Stock:");
            cells[startRow + 1, (byte)(startColumn + 1)].SetStyle(style);
        }


    }
}
