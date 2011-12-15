using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;


namespace Aspose.Cells.Demos
{
    /// <summary>
    /// Summary description for SchoolReport.
    /// </summary>
    public class SchoolReport : System.Web.UI.Page
    {
        protected System.Web.UI.WebControls.Button Button1;
        protected System.Web.UI.WebControls.ListBox ListBox1;
        protected System.Web.UI.WebControls.DropDownList ddlFileVersion;

        private void Page_Load(object sender, System.EventArgs e)
        {
            // Put user code to initialize the page here

            if (!IsPostBack)
            {
                CreateList();
            }
            else
            {
                if (this.Request.Params.Count > 0)
                {
                    string param = this.Request.Params[0];
                    //Instantiate a workbook
                    Workbook workbook = new Workbook();
                    CreateStaticReport(workbook);
                    CreateDynamicReport(workbook);


                    if (ddlFileVersion.SelectedItem.Value == "XLS")
                    {
                        ////Save file and send to client browser using selected format
                        workbook.Save(HttpContext.Current.Response, "ReportCard.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
                    }
                    else
                    {
                        workbook.Save(HttpContext.Current.Response, "ReportCard.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
                    }

                    //end response to avoid unneeded html
                    HttpContext.Current.Response.End();   
                }

            }

        }

        #region Web Form Designer generated code
        override protected void OnInit(EventArgs e)
        {
            //
            // CODEGEN: This call is required by the ASP.NET Web Form Designer.
            //
            InitializeComponent();
            base.OnInit(e);
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Button1.Click += new System.EventHandler(this.Button1_Click);
            this.Load += new System.EventHandler(this.Page_Load);

        }
        #endregion

        /// <summary>
        /// Creates student list.
        /// </summary>
        private void CreateList()
        {
            //Creates student list from data in an Workbook file. 
            //In a real world application, all kind of data sources can be used.
           
            //Open the template
            string path = System.Web.HttpContext.Current.Server.MapPath("~");
            path = path.Substring(0, path.LastIndexOf("\\"));
            path += @"\designer\SchoolData.xls";


            //string path = MapPath("~/designer/SchoolData.xls");

            string dataFile = path;
            Workbook workbook = new Workbook(dataFile);
 

            //Get the cells collection in the first worksheet
            Cells cells = workbook.Worksheets[0].Cells;
            //Export the sheet data to a multi-dimensional array
            object[,] nameList = cells.ExportArray(1, 0, cells.MaxDataRow, 2);
            //Fill the list box
            for (int i = 0; i < nameList.Length / 2; i++)
            {
                this.ListBox1.Items.Add(nameList[i, 0].ToString() + " " + nameList[i, 1].ToString());
            }
            //Set first element as a selected item
            this.ListBox1.SelectedIndex = 0;
        }

        private void Button1_Click(object sender, System.EventArgs e)
        {
            //Redirect to the same page with some parameter
            Response.Redirect("SchoolReport.aspx?Data=abc");
        }

        private void CreateStaticReport(Workbook workbook)
        {
            //Sets default font
            Style style = workbook.DefaultStyle;
            style.Font.Name = "Tahoma";
            workbook.DefaultStyle = style;

            //Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];
            //Name the worksheet
            sheet.Name = "Report Card";
            //Make the gridlines insible for the sheet
            sheet.IsGridlinesVisible = false;

            AddImageAndChart(workbook);

            //Add a new worksheet to the workbook
            int index = workbook.Worksheets.Add();
            //Get the sheet
            sheet = workbook.Worksheets[index];
            //Name the sheet
            sheet.Name = "Grade Table";
            //Make the gridlines invisible for the worksheet
            sheet.IsGridlinesVisible = false;


            SetRowColumn(workbook);
            CreateOutline(workbook);
            CreateCellsFormatting(workbook);
            CreateStaticData(workbook);



        }

        private void SetRowColumn(Workbook workbook)
        {
            //Get the cells in the first worksheet
            Cells cells = workbook.Worksheets[0].Cells;
            //Set the height of the first 22 rows
            for (int i = 0; i < 22; i++)
                cells.SetRowHeight(i, 13.5);

            //Set the height of the next row
            cells.SetRowHeight(22, 6);
            //Set the height for the 24th row
            cells.SetRowHeight(23, 13.5);
            //Set the row height for the next 5 rows
            for (int i = 24; i < 29; i++)
                cells.SetRowHeight(i, 22.5);
            //Set the row height for the next two rows
            cells.SetRowHeight(29, 13.5);
            cells.SetRowHeight(30, 6);

            //Set the row height for the 32-34 rows
            for (int i = 31; i < 34; i++)
                cells.SetRowHeight(i, 13.5);


            //Set the columns widths for first four (A-D) columns
            cells.SetColumnWidth(0, 1.86);
            cells.SetColumnWidth(1, 1.86);
            cells.SetColumnWidth(2, 19);
            cells.SetColumnWidth(3, 15.14);

            //Set the column widths for 5-10 columns
            for (byte column = 4; column < 10; column++)
                cells.SetColumnWidth(column, 5);

            //Set the column widths for 12-13 columns
            cells.SetColumnWidth(11, 15.43);
            cells.SetColumnWidth(12, 2);

            //Get the third worksheet cells
            cells = workbook.Worksheets[2].Cells;
            //Set the row height for the second row
            cells.SetRowHeight(1, 15.75);
            //Set the column widths for first two columns
            cells.SetColumnWidth(0, 2);
            cells.SetColumnWidth(1, 11.86);
            //Set the Column widths for 3-15 columns
            for (int i = 2; i < 15; i++)
                cells.SetColumnWidth(i, 5.14);
        }

        private void CreateOutline(Workbook workbook)
        {
            //Get the first worksheet cells
            Cells cells = workbook.Worksheets[0].Cells;

            //Create a range and its outline borders
            Range range = cells.CreateRange("B2", "M22");
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));
            range.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));

            //Create a range and its outline borders
            range = cells.CreateRange("B24", "M30");
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));
            range.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));

            //Create a range and its outline borders
            range = cells.CreateRange("B32", "M34");
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));
            range.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128));
        }

        private void CreateCellsFormatting(Workbook workbook)
        {
            //Creates cell formatting on the first worksheet

            //Create a style object
            Style style = workbook.Styles[workbook.Styles.Add()];
            //Sets font attributes
            style.Font.IsBold = true;
            style.Font.Size = 10;

            //Set the style to some cells
            Cells cells = workbook.Worksheets[0].Cells;
            cells["K4"].SetStyle(style);
            cells["E11"].SetStyle(style);
            cells["E12"].SetStyle(style);

            //Create the style object
            style = workbook.Styles[workbook.Styles.Add()];
            //Sets font attributes
            style.Font.IsBold = true;
            style.HorizontalAlignment = TextAlignmentType.Center;

            //Sets borders
            style.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            //Apply style to some cells (C15-L15)
            int startColumn = CellsHelper.ColumnNameToIndex("C");
            int endColumn = CellsHelper.ColumnNameToIndex("L");
            for (int i = startColumn; i <= endColumn; i++)
            {
                cells[14, i].SetStyle(style);
            }

            //Create the style object and set borders
            style = workbook.Styles[workbook.Styles.Add()];
            style.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            //Sets foreground color
            style.ForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xCC);
            style.Pattern = BackgroundType.Solid;

            //Apply style to some specific cells in rows
            for (int i = startColumn; i <= endColumn; i++)
            {
                cells[15, i].SetStyle(style);
                cells[17, i].SetStyle(style);
                cells[19, i].SetStyle(style);
            }

            //Create the style and set borders
            style = workbook.Styles[workbook.Styles.Add()];
            style.Borders[BorderType.TopBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            //Sets foreground color
            style.ForegroundColor = Color.FromArgb(0xCC, 0xFF, 0xCC);
            style.Pattern = BackgroundType.Solid;

            //Apply style to some cells in rows
            for (int i = startColumn; i <= endColumn; i++)
            {
                cells[16, i].SetStyle(style);
                cells[18, i].SetStyle(style);
                cells[20, i].SetStyle(style);
            }

            //Create the style object and set borders
            style = workbook.Styles[workbook.Styles.Add()];
            style.Borders[BorderType.BottomBorder].Color = Color.FromArgb(0, 0, 128);
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

            //Apply style to some cells in rows
            for (int i = startColumn; i <= endColumn; i++)
            {
                cells[24, i].SetStyle(style);
                cells[25, i].SetStyle(style);
                cells[26, i].SetStyle(style);
                cells[27, i].SetStyle(style);
                cells[28, i].SetStyle(style);
            }

            //Apply style to some cells in a row
            for (int i = 4; i <= 9; i++)
                cells[32, i].SetStyle(style);
            cells[32, 11].SetStyle(style);

            //Apply some custom number style in a column's cells
            for (int i = 15; i < 21; i++)
            {
                Aspose.Cells.Style style1 = new Style();
                style.Custom = "0";
                cells[i, 10].SetStyle(style1);
            }

            //Get the second worksheet cells collection
            cells = workbook.Worksheets[2].Cells;

            //Create the style object
            style = workbook.Styles[workbook.Styles.Add()];
            //Specify the forground color and font attributes
            style.ForegroundColor = Color.FromArgb(128, 0, 0);
            style.Pattern = BackgroundType.Solid;
            style.Font.Color = Color.FromArgb(255, 255, 153);
            style.Font.Size = 12;
            style.Font.IsBold = true;

            //Apply the style to some cells in the second row
            for (int i = 1; i < 15; i++)
                cells[1, i].SetStyle(style);

            //Create the style object and set borders
            style = workbook.Styles[workbook.Styles.Add()];
            style.Borders[BorderType.TopBorder].Color = Color.Black;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = Color.Black;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].Color = Color.Black;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].Color = Color.Black;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            //Specify the alignment type
            style.HorizontalAlignment = TextAlignmentType.Center;
            //Set the font attribute
            style.Font.IsBold = true;
            //Apply style to B3 and B4 cells
            cells["B3"].SetStyle(style);
            cells["B4"].SetStyle(style);

            //Create the style and set borders
            style = workbook.Styles[workbook.Styles.Add()];
            style.Borders[BorderType.TopBorder].Color = Color.Black;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = Color.Black;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].Color = Color.Black;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].Color = Color.Black;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            //Set alignment type
            style.HorizontalAlignment = TextAlignmentType.Center;
            //Apply style to some range of cells
            for (int i = 2; i < 4; i++)
            {
                for (int j = 2; j < 15; j++)
                    cells[i, j].SetStyle(style);
            }
        }

        private void CreateStaticData(Workbook workbook)
        {
            //Get the cells in the first worksheet
            Cells cells = workbook.Worksheets[0].Cells;

            //Get Style Object 
            Aspose.Cells.Style style = cells["E4"].GetStyle();

            //Put values, apply formula(s) to the cells with style formatting
            cells["E4"].PutValue("ASPOSE School District");
            style.Font.IsBold = true;
            style.Font.Size = 12;
            cells["E4"].SetStyle(style);
            cells["K4"].PutValue("Progress Report");
            cells["E5"].PutValue("Suite 180, 9 Crofts Avenue");
            cells["K5"].PutValue("Date:");
            style.Font.IsBold = true;
            cells["K5"].SetStyle(style);
            cells["L5"].Formula = "=Now()";
            style.Custom = "[$-409]mmmm d, yyyy;@";
            cells["L5"].SetStyle(style);
            cells["E6"].PutValue("Hurstville, NSW, 2220");
            cells["E8"].PutValue("Phone: 888.277.6734");
            cells["E9"].PutValue("Fax: 866.810.9465");
            cells["E11"].PutValue("Student Name");
            cells["E12"].PutValue("Student SSN");
            cells["C15"].PutValue("Class Name");
            cells["D15"].PutValue("Teacher");
            cells["E15"].PutValue("1st");
            cells["F15"].PutValue("2nd");
            cells["G15"].PutValue("3rd");
            cells["H15"].PutValue("4th");
            cells["I15"].PutValue("5th");
            cells["J15"].PutValue("6th");
            cells["K15"].PutValue("Final");
            cells["L15"].PutValue("Letter Grade");
            cells["C16"].PutValue("English");
            cells["C17"].PutValue("Math");
            cells["C18"].PutValue("Social Studies");
            cells["C19"].PutValue("Science");
            cells["C20"].PutValue("Art");
            cells["C21"].PutValue("Physical Education");
            cells["C24"].PutValue("Note");
            style.Font.IsBold = true;
            cells["C24"].SetStyle(style);
            cells["D33"].PutValue("Parent Signature:");
            cells["D33"].SetStyle(style);
            cells["K33"].PutValue("Date");
            cells["K33"].SetStyle(style);
            cells = workbook.Worksheets[2].Cells;
            cells["B2"].PutValue("Grade Table");
            cells["B3"].PutValue("Average");
            cells["C3"].PutValue(0);
            cells["D3"].PutValue(60);
            cells["E3"].PutValue(63);
            cells["F3"].PutValue(67);
            cells["G3"].PutValue(70);
            cells["H3"].PutValue(73);
            cells["I3"].PutValue(77);
            cells["J3"].PutValue(80);
            cells["K3"].PutValue(83);
            cells["L3"].PutValue(87);
            cells["M3"].PutValue(90);
            cells["N3"].PutValue(93);
            cells["O3"].PutValue(97);
            cells["B4"].PutValue("Letter Grade");
            cells["C4"].PutValue("F");
            cells["D4"].PutValue("D-");
            cells["E4"].PutValue("D");
            cells["F4"].PutValue("D+");
            cells["G4"].PutValue("C-");
            cells["H4"].PutValue("C");
            cells["I4"].PutValue("C+");
            cells["J4"].PutValue("B-");
            cells["K4"].PutValue("B");
            cells["L4"].PutValue("B+");
            cells["M4"].PutValue("A-");
            cells["N4"].PutValue("A");
            cells["O4"].PutValue("A+");

        }

        private void AddImageAndChart(Workbook workbook)
        {
            //Get the image file path
            string path = System.Web.HttpContext.Current.Server.MapPath("~");
            path = path.Substring(0, path.LastIndexOf("\\"));
            path += @"\Image\School.jpg";


            string imageFile = path;
            //Add image to the first worksheet
            int index = workbook.Worksheets[0].Pictures.Add(1, 1, imageFile);
            Picture pic = workbook.Worksheets[0].Pictures[index];
            pic.Left = 2;
            pic.Top = 2;

            //Add a chart worksheet type
            index = workbook.Worksheets.Add(SheetType.Chart);
            //Get the worksheet
            Worksheet sheet = workbook.Worksheets[index];
            //Set the name
            sheet.Name = "Grade Chart";
            //Set the scalling factor
            sheet.Zoom = 90;

            
            //Add a new bar chart to the worksheet            
            Chart chart = sheet.Charts[sheet.Charts.Add(ChartType.Bar, 0, 0, 0, 0)];

            //Set the nseries data range
            chart.NSeries.Add("'Report Card'!E16:H21", false);
            //Name the series
            for (int i = 0; i < chart.NSeries.Count; i++)
                chart.NSeries[i].Name = "='Report Card'!C" + (16 + i).ToString();

            //Set the legend position to bottom on the chart
            chart.Legend.Position = LegendPositionType.Bottom;
        }

        private void CreateDynamicReport(Workbook workbook)
        {
            //Get the template file path
            string path = MapPath("~");
            path = path.Substring(0, path.LastIndexOf("\\"));
            string dataFile = path + "\\Designer\\SchoolData.xls";
            //Get the selected list box item
            string name = this.ListBox1.SelectedItem.Text;
            //Split the array
            string[] nameArray = name.Split(' ');
            //Get the first worksheet cells
            Cells cells = workbook.Worksheets[0].Cells;
            //Put the selected value (in the list box) to the cell
            cells["H11"].PutValue(name);

            //Instantiate a workbook
            Workbook dataWorkbook = new Workbook(dataFile);

            //Get teachers' name
            string[] teachers = new string[dataWorkbook.Worksheets.Count];
            for (int i = 0; i < teachers.Length; i++)
                teachers[i] = dataWorkbook.Worksheets[i].Cells["L1"].StringValue;

            //Put teachers' name into output workbook
            cells.ImportArray(teachers, 15, 3, true);


            //Get / Set students data
            Cell cell = null;
            Worksheet dataSheet = dataWorkbook.Worksheets[dataWorkbook.Worksheets.Count - 1];
            for (; ; )
            {
                cell = dataSheet.Cells.FindString(nameArray[0], cell);
                if (cell != null)
                {
                    if (dataSheet.Cells[cell.Row, cell.Column + 1].StringValue == nameArray[1])
                    {
                        cells["H12"].PutValue(dataSheet.Cells[cell.Row, cell.Column + 2].Value);
                        break;
                    }
                }
                else
                    break;
            }

            for (int i = 0; i < dataWorkbook.Worksheets.Count - 1; i++)
            {
                DataTable studentData = dataWorkbook.Worksheets[i].Cells.ExportDataTable(1, 0, cells.MaxDataRow, 8);
                foreach (DataRow row in studentData.Rows)
                {
                    if (row[0].ToString() == nameArray[0] && row[1].ToString() == nameArray[1])
                    {
                        for (int j = 2; j < row.ItemArray.Length; j++)
                        {
                            cells[15 + i, j + 2].PutValue(row[j]);
                        }
                    }
                }
            }

            //Specify some formulas for Marks Average and Grade
            for (int i = 15; i < 21; i++)
            {
                cells[i, 10].Formula = "=AVERAGE(E" + (i + 1).ToString() + ":J" + (i + 1).ToString() + ")";
                cells[i, 11].Formula = "=IF(K" + (i + 1).ToString() + "<>\"\",HLOOKUP(K"
                    + (i + 1).ToString() + ",'Grade Table'!$C$3:$O$4,2),\"\")";
            }
            
            //Calculate all the formulas
            workbook.CalculateFormula();

            int courseIndex = -1;
            double minScore = -1;
            for (int i = 15; i < 21; i++)
            {
                double score = cells[i, 10].DoubleValue;
                if (score < 80)
                {
                    if (minScore < 0)
                    {
                        minScore = score;
                        courseIndex = i;
                    }
                    else if (score < minScore)
                    {
                        minScore = score;
                        courseIndex = i;
                    }
                }
            }

            //Specify some notes for the marksheet
            if (courseIndex != -1)
            {
                string course = cells[courseIndex, 2].StringValue;
                string note = "{0} seems to be having difficulties with {1} projects.  We offer after school tutoring sessions"; ;
                note = string.Format(note, nameArray[0], course);
                cells["C25"].PutValue(note);
                cells["C26"].PutValue("which may be helpful.");
            }

        }
    }
}


