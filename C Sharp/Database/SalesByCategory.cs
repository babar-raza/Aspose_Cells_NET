using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for SalesByCategory.
    /// </summary>
    public class SalesByCategory : DbBase
    {
        public SalesByCategory(string path)
            : base(path)
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public Workbook CreateSalesByCategory()
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
                this.oleDbDataAdapter1.SelectCommand.CommandText = @"SELECT DISTINCTROW Categories.CategoryID, 
					Categories.CategoryName, Products.ProductName, SUM([Order Details Extended].ExtendedPrice) AS ProductSales
				FROM  Categories 
				INNER JOIN
					(Products INNER JOIN (Orders INNER JOIN [Order Details Extended] ON
					Orders.OrderID = [Order Details Extended].OrderID) ON Products.ProductID = [Order Details Extended].ProductID) ON Categories.CategoryID = Products.CategoryID
				WHERE
					(((Orders.OrderDate) BETWEEN #1/1/1995# AND 
						#12/31/1995#)) GROUP BY Categories.CategoryID ,  Categories.CategoryName ,  Products.ProductName ORDER BY Categories.CategoryName";
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
            Worksheet sheet = workbook.Worksheets["Sheet8"];
            //Name the worksheet
            sheet.Name = "Sales By Category";
            //Get the cells
            Cells cells = sheet.Cells;
            //Get the vertical page breaks
            VerticalPageBreakCollection vPageBreaks = sheet.VerticalPageBreaks;
            int currentRow = 2;
            byte currentColumn = 0;

            string lastCategory = "";
            string thisCategory, nextCategory;

            SetSalesByCategoryStyles(workbook);
            //Fill cells with source data and apply styles
            for (int i = 0; i < this.dataTable1.Rows.Count; i++)
            {
                thisCategory = (string)this.dataTable1.Rows[i]["CategoryName"];
                if (thisCategory != lastCategory)
                {
                    currentRow = 2;
                    if (i != 0)
                        currentColumn += 15;
                    CreateSalesByCategoryHeader(workbook, cells, currentRow, currentColumn, thisCategory);
                    lastCategory = thisCategory;
                    currentRow += 2;
                }
                cells[currentRow, currentColumn].PutValue((string)this.dataTable1.Rows[i]["ProductName"]);
                cells[currentRow, (byte)(currentColumn + 1)].PutValue((double)(decimal)this.dataTable1.Rows[i]["ProductSales"]);

                cells[currentRow, (byte)(currentColumn + 1)].SetStyle(workbook.Styles["Sales"]);

                cells.SetColumnWidth(currentColumn, 27);
                cells.SetColumnWidth((byte)(currentColumn + 1), 15);

                if (i != this.dataTable1.Rows.Count - 1)
                {
                    nextCategory = (string)this.dataTable1.Rows[i + 1]["CategoryName"];
                    if (thisCategory != nextCategory)
                    {
                        vPageBreaks.Add(0, currentColumn + 1);
                        CreateChart(workbook, sheet, currentRow, currentColumn);
                    }
                }
                else
                {
                    CreateChart(workbook, sheet, currentRow, currentColumn);
                }
                currentRow++;
            }
            //Remove the unnecessary worksheets
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                sheet = workbook.Worksheets[i];
                if (sheet.Name != "Sales By Category")
                {
                    workbook.Worksheets.RemoveAt(i);
                    i--;
                }
            }
            //Get the workbook (generated)
            return workbook;
        }

        private void CreateChart(Workbook workbook, Worksheet sheet, int currentRow, int currentColumn)
        {
            //Add a bar chart
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 4, currentColumn + 3,
                26, currentColumn + 14);
            //Get the chart
            Chart chart = sheet.Charts[chartIndex];
            //Make the legends invisible
            chart.ShowLegend = false;
            string startCell = CellsHelper.CellIndexToName(4, currentColumn + 1);
            string endCell = CellsHelper.CellIndexToName(currentRow, currentColumn + 1);
            //Set the nseries
            chart.NSeries.Add(startCell + ":" + endCell, true);
            //Set the fill format for the plot area
            FillFormat fillFormat = chart.PlotArea.Area.FillFormat;
            fillFormat.SetPresetColorGradient(GradientPresetType.Daybreak, GradientStyleType.Vertical, 1);
            //Set the category data
            startCell = CellsHelper.CellIndexToName(4, currentColumn);
            endCell = CellsHelper.CellIndexToName(currentRow, currentColumn);
            chart.NSeries.CategoryData = startCell + ":" + endCell;
        }
        private void CreateSalesByCategoryHeader(Workbook workbook, Cells cells, int currentRow, byte currentColumn, string categoryName)
        {
            //Input data and apply style
            cells[currentRow, currentColumn].PutValue(categoryName);
            cells[currentRow, currentColumn].SetStyle( workbook.Styles["Header"]);
        }

        private void SetSalesByCategoryStyles(Workbook workbook)
        {
            //Create a style with specific formatting attributes
            int styleIndex = workbook.Styles.Add();
            Style style = workbook.Styles[styleIndex];
            style.Number = 7;
            style.Name = "Sales";

            //Create a style withe specific set of attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Font.Size = 14;
            style.Font.IsBold = true;
            style.Font.IsItalic = true;
            style.Font.Color = Color.Yellow;
            style.ForegroundColor = Color.Blue;
            style.Pattern = BackgroundType.Solid;
            style.Name = "Header";
        }

    }
}


