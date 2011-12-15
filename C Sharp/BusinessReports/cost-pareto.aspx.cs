using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;


namespace Aspose.Cells.Demos
{
    public partial class CostPareto : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnProcess_Click(object sender, EventArgs e)
        {            
            //Create a dataset object
            DataSet ds = new DataSet();
            
            //Get data from xml file
            string path = MapPath(".");
            path = path.Substring(0, path.LastIndexOf("\\"));
            path += @"\Database\CostPareto.xml";
            
            //Load data from xml file to dataset
            ds.ReadXml(path, XmlReadMode.ReadSchema);

            //Create a new workbook
            Workbook workbook = new Workbook();
            
            //Generate first data sheet
            GenerateDataSheet(workbook, ds);
            
            //Generate second chart sheet
            GenerateChartSheet(workbook, ds);

            //Create an object of SaveFormat
            SaveFormat saveFormat = new SaveFormat();
            
            //Check file format is xls
            if(ddlFileVersion.SelectedItem.Value == "XLS")
            {
                //Set save format optoin to xls
                saveFormat = SaveFormat.Excel97To2003;
            }
            //Check file format is xlsx
            else if (ddlFileVersion.SelectedItem.Value == "XLSX")
            {
                //Set save format optoin to xlsx
                saveFormat = SaveFormat.Xlsx;                
            }

            //Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "CostPareto." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));
			
            // note by Vit - end response to avoid unneeded html after xls
            Response.End();
        }


        private void GenerateDataSheet(Workbook workbook, DataSet ds)
        {
            //Write data to first data sheet
            Worksheet sheet1 = workbook.Worksheets[0];
            
            //Name the sheet
            sheet1.Name = "Cost Data";
            
            //Write sheet1 cells data to cells object
            Cells cells = sheet1.Cells;

            //Import data into cells
            cells.ImportDataTable(ds.Tables[0], true, 0, 0, ds.Tables[0].Rows.Count, ds.Tables[0].Columns.Count);

            //Set header style with specific formatting attributes
            StyleCollection styles = workbook.Styles;
            
            //Set style index
            int styleIndex = styles.Add();

            //Set style attribute using style index
            Style style = styles[styleIndex];

            //Set font size 
            style.Font.Size = 10;

            //Set font color to white
            style.Font.Color = Color.White;

            //Set font to bold
            style.Font.IsBold = true;

            //Set font name to Verdana
            style.Font.Name = "Verdana";

            //Locked style
            style.IsLocked = true;

            //Set vertical alignment 
            style.VerticalAlignment = TextAlignmentType.Center;

            //Set horizontal alignment
            style.HorizontalAlignment = TextAlignmentType.Left;

            //Set indent level
            style.IndentLevel = 1;

            //Set top, bottom, left and right borders style
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            //Change the palette for the spreadsheet in the specified index
            workbook.ChangePalette(Color.FromArgb(10, 100, 180), 50);

            //Change foreground color
            style.ForegroundColor = Color.FromArgb(10, 100, 180);

            //Set background style pattern
            style.Pattern = BackgroundType.Solid;

            //Set first two column's widths and set the height of the first row
            cells.SetColumnWidth(0, 25);
            cells.SetColumnWidth(1, 18);
            cells.SetRowHeight(0, 30);
            
            //Apply the style to A1 cell
            cells[0, 0].SetStyle(style);

            //Add a new style
            styleIndex = styles.Add();
            Style style1 = styles[styleIndex];
            
            //Copy above created style to it
            style1.Copy(style);

            //Set horizontal alignment and indentation
            style1.HorizontalAlignment = TextAlignmentType.Right;
            style1.IndentLevel = 0;
            
            //Apply the style to B1 cell
            cells[0, 1].SetStyle(style1);

            //Set current row to 1
            int currentRow = 1;
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                //Set row height and color
                cells.SetRowHeight(currentRow, 20);
                Color color = Color.FromArgb(255, 255, 255);

                //Change palette color of workbook
                workbook.ChangePalette(color, 51);

                //Change color of even number rows
                if (currentRow % 2 == 0)
                {
                    //Set color
                    color = Color.FromArgb(250, 250, 200);

                    //Change palette color of workbook
                    workbook.ChangePalette(color, 52);
                }

                //Set style for the first column cells
                styleIndex = styles.Add();

                //Set style attribute using style index
                Style styleCell1 = styles[styleIndex];

                //Set font size
                styleCell1.Font.Size = 10;

                //Set font name 
                styleCell1.Font.Name = "Arial";

                //Set horizontal alignment
                styleCell1.HorizontalAlignment = TextAlignmentType.Left;

                //Set vertical alignment
                styleCell1.VerticalAlignment = TextAlignmentType.Center;

                //Set indenting level
                styleCell1.IndentLevel = 1;

                //Set top, bottom, left and right borders style
                styleCell1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleCell1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Dashed;
                styleCell1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.None;
                styleCell1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.None;

                //Check for last row
                if (currentRow == ds.Tables[0].Rows.Count)
                {
                    //Set bottom border style of last row
                    styleCell1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                }

                //Set foreground color
                styleCell1.ForegroundColor = color;

                //Set background pattern style
                styleCell1.Pattern = BackgroundType.Solid;

                //Apply style to current row in first column
                cells[currentRow, 0].SetStyle(styleCell1);

                //Set style for the second column cells
                styleIndex = styles.Add();

                //Set style attribute using style index
                Style styleCell2 = styles[styleIndex];

                //Copy previous style in new attribute
                styleCell2.Copy(styleCell1);

                //Set left and right border style
                styleCell2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.None;
                styleCell2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                //Set horizontal text alignment 
                styleCell2.HorizontalAlignment = TextAlignmentType.Right;

                //Set indent level
                styleCell2.IndentLevel = 0;

                //Set number format
                styleCell2.Custom = "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)";
                
                //Apply style to current row in second column
                cells[currentRow, 1].SetStyle(styleCell2);

                //Add 1 in current row count
                currentRow++;
            }
        }

        private void GenerateChartSheet(Workbook workbook, DataSet ds)
        {
            //Generate the second chart sheet
            int sheetIndex = workbook.Worksheets.Add(SheetType.Chart);
            Worksheet sheet2 = workbook.Worksheets[sheetIndex];
           
            //Name the sheet
            sheet2.Name = "Pareto Chart";
            
            //Set chart index
            int chartIndex = sheet2.Charts.Add(ChartType.Column, 0, 0, 0, 0);
            
            //Get chart type
            Chart chart = sheet2.Charts[chartIndex];

            //Set chart title text
            chart.Title.Text = "Cost Center";

            //Set chart title font
            chart.Title.TextFont.IsBold = true;
            chart.Title.TextFont.Size = 16;

            //Set series
            string series = "Cost Data!B2:B" + (ds.Tables[0].Rows.Count + 1);

            //Series add in chart
            chart.NSeries.Add(series, true);

            //Set series name
            chart.NSeries[0].Name = "Annual Cost";
           
            //Set category
            chart.NSeries.CategoryData = "Cost Data!A2:A" + (ds.Tables[0].Rows.Count + 1);
            
            //Legend not shown
            chart.ShowLegend = false;

            //Set chart style
            workbook.ChangePalette(Color.FromArgb(255, 255, 200), 53);

            //Set plot area foreground color
            chart.PlotArea.Area.ForegroundColor = Color.FromArgb(255, 255, 200);
            
            //Set major grid line color
            workbook.ChangePalette(Color.FromArgb(121, 117, 200), 54);
            chart.CategoryAxis.MajorGridLines.Color = Color.FromArgb(121, 117, 200);

            //Set series each point color
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                workbook.ChangePalette(Color.FromArgb(10, 100, 180), 55);
                chart.NSeries[0].Points[i].Area.ForegroundColor = Color.FromArgb(10, 100, 180);
                workbook.ChangePalette(Color.FromArgb(255, 255, 200), 53);
                chart.NSeries[0].Points[i].Border.Color = Color.FromArgb(255, 255, 200);
            }
        }


    }
}


