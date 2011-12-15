using System;
using System.Data;
using System.Drawing;
using Aspose.Cells.Drawing;
namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for Invoice.
    /// </summary>
    public class Invoice : DbBase
    {
        public Invoice(string path)
            : base(path)
        {

        }

        public Workbook CreateInvoice()
        {
            try
            {
                DBInit();

                //Specify SQL for command text
                this.oleDbDataAdapter1.SelectCommand.CommandText = "SELECT DISTINCTROW OrderID FROM Orders ORDER BY OrderID DESC";
                //Fill the datatable 
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

            DataTable[] dtInvoice = new DataTable[this.dataTable1.Rows.Count];

            //for(int i = 0; i < dtInvoice.Length; i ++)
            //We generate invoices for the first 100 orders for demo only. If you want to
            //generate all invoices,uncomment the line above and comment the line below.
            for (int i = 0; i < 50; i++)
                dtInvoice[i] = this.ReadInvoice(this.dataTable1.Rows[i][0].ToString());

            //Create the workbook
            Workbook workbook = new Workbook();
            //Get all the worksheets
            WorksheetCollection sheets = workbook.Worksheets;
            //get the first worksheet
            Worksheet sheet = sheets[0];
            //Name the worksheet
            sheet.Name = "Invoice";
            //Get the sheet cells
            Cells cells = sheet.Cells;
            int startRow = 0;

            SetInvoiceStyles(workbook);
            string imagePath = path + "\\Image";
            //for(int i = 0; i < dtInvoice.Length; i ++)
            //We generate invoices for the first 100 orders for demo only. If you want to
            //generate all invoices,uncomment the line above and comment the line below.
            for (int i = 0; i < 50; i++)
            {
                //Add picture(s)
                sheet.Pictures.Add(startRow, 0, startRow + 2, 1, imagePath + "\\logo.jpg");
                int picIndex = sheet.Pictures.Add(startRow, 1, startRow + 2, 2, imagePath + "\\namelogo.jpg");
                Picture pic = sheet.Pictures[picIndex];
                pic.UpperDeltaY = 100;

                CreateInvoiceHeader(cells, workbook, dtInvoice[i], startRow);
                startRow += 11;
                CreateOrder(cells, workbook, dtInvoice[i], startRow, this.dataTable1.Rows[i][0].ToString());
                startRow += 4;
                CreateOrderDetail(cells, workbook, dtInvoice[i], startRow);

                startRow += dtInvoice[i].Rows.Count + 1;
                //Add horizontal page break(s)
                sheet.HorizontalPageBreaks.Add(startRow - 1, 0);
            }

            //Get the workbook (generated)
            return workbook;

        }
        private DataTable ReadInvoice(string orderID)
        {
            try
            {
                //Specify SQL
                string invoiceQuery = "SELECT DISTINCTROW Invoices.* FROM Invoices WHERE Invoices.OrderID="
                    + orderID;
                //Specify the command
                this.oleDbDataAdapter2.SelectCommand.CommandText = invoiceQuery;
            }
            catch
            {
            }
            finally
            {
                if (this.oleDbConnection1 != null)
                    this.oleDbConnection1.Close();
            }
            //Create a datatable and fill it with data based on the query
            DataTable dtInvoice = new DataTable();
            this.oleDbDataAdapter2.Fill(dtInvoice);
            //Retrieve the datatable
            return dtInvoice;
        }

        private void SetInvoiceStyles(Workbook workbook)
        {
            //Add LightBlue and DarkBlue colors to color palette
            workbook.ChangePalette(Color.LightBlue, 54);
            workbook.ChangePalette(Color.DarkBlue, 55);
            
            //Create a style with specific formatting attributes
            Style style;
            int styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.Font.Color = Color.White;
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.Name = "Font12Center";

            //Create a style with specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.Font.Color = Color.White;
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Left;
            style.Name = "Font12Left";

            //Create a style with specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.Font.Color = Color.White;
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Right;
            style.Name = "Font12Right";

            //Create a style with specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Number = 7;
            style.Name = "Number7";

            //Create a style with specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Number = 9;
            style.Name = "Number9";

            //Create a style with specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.Name = "Center";

            //Create a style with specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Font.Size = 16;
            style.Font.IsBold = true;
            style.Font.Color = Color.DarkBlue;
            style.Name = "Darkblue";

            //Create a style with specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.Font.Color = Color.DarkBlue;
            style.Name = "Darkblue12";

            //Create a style with specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Font.IsItalic = true;
            style.Font.Color = Color.DarkBlue;
            style.Name = "DarkblueItalic";

            //Create a style with specific formatting attributes
            styleIndex = workbook.Styles.Add();
            style = workbook.Styles[styleIndex];
            style.Borders[BorderType.BottomBorder].Color = Color.Black;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Medium;
            style.Name = "BlackMedium";


        }

        private void CreateOrderDetail(Cells cells, Workbook workbook, DataTable dtInvoice, int startRow)
        {
            //Define some styles
            Style style1, style2, style3;
            //Get styles
            style1 = workbook.Styles["Number7"];
            style2 = workbook.Styles["Number9"];
            style3 = workbook.Styles["Center"];

            //Fill cells based on differnt datatable fields
            //Apply styles to cells too
            for (int i = 0; i < dtInvoice.Rows.Count; i++)
            {
                cells[startRow + i, 0].PutValue((int)dtInvoice.Rows[i]["ProductID"]);
                cells[startRow + i, 0].SetStyle( style3);
                cells[startRow + i, 1].PutValue((string)dtInvoice.Rows[i]["ProductName"]);
                cells[startRow + i, 3].PutValue((short)dtInvoice.Rows[i]["Quantity"]);
                cells[startRow + i, 4].PutValue((double)(decimal)dtInvoice.Rows[i]["UnitPrice"]);
                cells[startRow + i, 4].SetStyle( style1);
                cells[startRow + i, 5].PutValue((float)dtInvoice.Rows[i]["Discount"]);
                cells[startRow + i, 5].SetStyle(style2);
                cells[startRow + i, 6].PutValue((double)(decimal)dtInvoice.Rows[i]["ExtendedPrice"]);
                cells[startRow + i, 6].SetStyle(style1);
            }
        }

        private void CreateOrder(Cells cells, Workbook workbook, DataTable dtInvoice, int startRow, string orderID)
        {
            //Set row heights for some rows
            cells.SetRowHeight(startRow, 14);
            cells.SetRowHeight(startRow + 3, 14);

            //Set column widths for some columns
            cells.SetColumnWidth(1, 16);
            cells.SetColumnWidth(2, 16);
            cells.SetColumnWidth(3, 16);
            cells.SetColumnWidth(4, 16);
            cells.SetColumnWidth(5, 16);
            cells.SetColumnWidth(6, 18);
            //Get the style
            Style style = workbook.Styles["Font12Center"];
            //Apply the style to the cells
            for (byte i = 0; i < 7; i++)
            {
                cells[startRow, i].SetStyle(style);
                cells[startRow + 3, i].SetStyle(style);
            }
            //Get the style
            style = workbook.Styles["Center"];
            //Apply the style to the cells
            for (byte i = 0; i < 7; i++)
                cells[startRow + 1, i].SetStyle(style);

            //Input values to the cells based on the datatable
            cells[startRow, 0].PutValue("Order ID:");
            cells[startRow + 1, 0].PutValue(int.Parse(orderID));
            cells[startRow, 1].PutValue("Customer ID:");
            cells[startRow + 1, 1].PutValue((string)dtInvoice.Rows[0]["CustomerID"]);
            cells[startRow, 2].PutValue("Salesperson:");
            cells[startRow + 1, 2].PutValue((string)dtInvoice.Rows[0]["Salesperson"]);
            cells[startRow, 3].PutValue("Order Date:");
            cells[startRow + 1, 3].PutValue(((DateTime)dtInvoice.Rows[0]["OrderDate"]).ToString("D"));
            cells[startRow, 4].PutValue("Required Date:");
            cells[startRow + 1, 4].PutValue(((DateTime)dtInvoice.Rows[0]["RequiredDate"]).ToString("D"));
            cells[startRow, 5].PutValue("Shipped Date:");
            if (dtInvoice.Rows[0]["ShippedDate"] != DBNull.Value)
                cells[startRow + 1, 5].PutValue(((DateTime)dtInvoice.Rows[0]["ShippedDate"]).ToString("D"));
            cells[startRow, 6].PutValue("Ship Via:");
            cells[startRow + 1, 6].PutValue((string)dtInvoice.Rows[0]["Shippers.CompanyName"]);

            cells[startRow + 3, 0].PutValue("Product ID:");
            cells[startRow + 3, 1].PutValue("Product");
            cells[startRow + 3, 2].PutValue(" Name:");
            cells[startRow + 3, 3].PutValue("Quantity:");
            cells[startRow + 3, 4].PutValue("Unit Price:");
            cells[startRow + 3, 5].PutValue("Discount:");
            cells[startRow + 3, 6].PutValue("Extended Price:");

            //Get the style and apply it to the cell(s)
            style = workbook.Styles["Font12Right"];
            cells[startRow + 3, 1].SetStyle(style);

            //Get the style and apply it to the cell(s)
            style = workbook.Styles["Font12Left"];
            cells[startRow + 3, 2].SetStyle(style);
        }

        private void CreateInvoiceHeader(Cells cells, Workbook workbook, DataTable dtInvoice, int startRow)
        {
            
            //Set row height and column width 
            cells.SetRowHeight(startRow, 24);
            cells.SetColumnWidth(0, 12);

            //Input a value and set its style
            cells[startRow, 5].PutValue("INVOICE");
            Style style = workbook.Styles["Darkblue"];
            cells[startRow, 5].SetStyle(style);

            //Get the style and apply it to the cells
            style = workbook.Styles["BlackMedium"];
            for (int i = 0; i < byte.MaxValue; i++)
                cells[startRow + 2, (byte)i].SetStyle(style);

            //Input some values to the cells
            cells[startRow + 3, 0].PutValue("One Portals Way, Twin Points WA 98156");
            cells[startRow + 4, 0].PutValue("Phone:1-206-555-1417 Fax:1-206");
            style = workbook.Styles["DarkblueItalic"];
            cells[startRow + 3, 0].SetStyle(style);
            cells[startRow + 4, 0].SetStyle(style);

            //Get the current date
            DateTime currentDate = DateTime.Today;
            string strTime = currentDate.ToString("D");
            //Input date
            cells[startRow + 3, 5].PutValue("Date:");
            cells[startRow + 3, 6].PutValue(strTime);

            //Input a value
            cells[startRow + 6, 0].PutValue("Ship To:");
            //Get the style
            style = workbook.Styles["Darkblue12"];
            //Apply the style to a cell
            cells[startRow + 6, 0].SetStyle(style);
            //Set the related row height
            cells.SetRowHeight(startRow + 6, 16);
            //Input a value and apply style to it
            cells[startRow + 6, 4].PutValue("Bill To:");
            cells[startRow + 6, 4].SetStyle(style);
            //Apply the style to a cell
            cells[startRow + 3, 5].SetStyle(style);
            //Input values
            if (dtInvoice.Rows[0][0] != DBNull.Value)
            {
                cells[startRow + 6, 1].PutValue((string)dtInvoice.Rows[0][0]);
                cells[startRow + 6, 5].PutValue((string)dtInvoice.Rows[0][0]);
            }
            if (dtInvoice.Rows[0][1] != DBNull.Value)
            {
                cells[startRow + 7, 1].PutValue((string)dtInvoice.Rows[0][1]);
                cells[startRow + 7, 5].PutValue((string)dtInvoice.Rows[0][1]);
            }

            string strDest = "";
            if (dtInvoice.Rows[0][2] != DBNull.Value)
                strDest += dtInvoice.Rows[0][2];

            if (dtInvoice.Rows[0][3] != DBNull.Value)
                strDest += " " + dtInvoice.Rows[0][3];

            if (dtInvoice.Rows[0][4] != DBNull.Value)
                strDest += " " + dtInvoice.Rows[0][4];

            strDest.TrimStart(' ');

            if (strDest != "")
            {
                cells[startRow + 8, 1].PutValue(strDest);
                cells[startRow + 8, 5].PutValue(strDest);
            }
            cells[startRow + 9, 1].PutValue((string)dtInvoice.Rows[0][5]);
            cells[startRow + 9, 5].PutValue((string)dtInvoice.Rows[0][5]);

        }

    }
}


