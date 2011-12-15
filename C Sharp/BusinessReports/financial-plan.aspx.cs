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

namespace Aspose.Cells.Demos
{
    public partial class FinancialPlan : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnProcess_Click(object sender, EventArgs e)
        {
            //Open xls file template
            string path = MapPath("~/designer/FinancialPlan.xls");
           
            //Create a new workbook
            Workbook workbook = new Workbook(path);
            
            //Add style settings
            AddStyles(workbook);

            //Fill "Model Inputs" sheet
            FillModelInputs(workbook);

            //Fill "Profit and Loss" sheet
            FillProfitLoss(workbook);

            //Fill "Balance Sheet" sheet
            FillBalanceSheet(workbook);

            //Fill "Cash Flow" sheet
            FillCashFlow(workbook);

            //Fill "Loan Payment Calculator" sheet
            FillLoanPaymentCalculator(workbook);

            //Create an object of SaveFormat
            SaveFormat saveFormat = new SaveFormat();

            //Check file format is xls
            if (ddlFileVersion.SelectedItem.Value == "XLS")
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
            workbook.Save(HttpContext.Current.Response, "FinancialPlan." + ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, new XlsSaveOptions(saveFormat));			
           
            // note by Vit - end response to avoid unneeded html after xls
            Response.End();
        }

        private void AddStyles(Workbook workbook)
        {
            //Define a style object
            Aspose.Cells.Style style = null;
            
            //Create different styles (28) with common attributes.
            for (int i = 0; i < 28; i++)
            {
                //Add style to workbook
                int index = workbook.Styles.Add();
                style = workbook.Styles[index];
                //Name style
                style.Name = "Custom_Style" + ((int)(i + 1)).ToString();
                //Set style Foreground Color
                style.ForegroundColor = Color.White;
                //Set style Pattern
                style.Pattern = BackgroundType.Solid;
                //Set style alignments
                style.HorizontalAlignment = TextAlignmentType.Center;
                style.VerticalAlignment = TextAlignmentType.Bottom;
                //Set Style font face name and size
                style.Font.Name = "Arial";
                style.Font.Size = 10;
            }

            //Customize each style

            //Custom_Style1
            style = workbook.Styles["Custom_Style1"];
            style.Font.IsBold = true;
            style.Number = 37;

            //Custom_Style2
            style = workbook.Styles["Custom_Style2"];
            style.Font.IsBold = true;
            style.Custom = "\"$\"#,##0.00";

            //Custom_Style3
            style = workbook.Styles["Custom_Style3"];
            style.Font.IsItalic = true;
            style.Custom = "\"$\"#,##0";
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;

            //Custom_Style4
            style = workbook.Styles["Custom_Style4"];
            style.Font.IsItalic = true;
            style.Font.Underline = FontUnderlineType.Single;
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;

            //Custom_Style5
            style = workbook.Styles["Custom_Style5"];
            style.Font.IsBold = true;
            style.Number = 10;	//0.00%

            //Custom_Style6
            style = workbook.Styles["Custom_Style6"];
            style.Font.IsBold = true;
            style.Number = 9;
            //Set style border line thickness
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

            //Custom_Style7
            style = workbook.Styles["Custom_Style7"];
            style.Font.IsBold = true;
            style.Number = 38;
            //Set style border line thickness
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

            //Custom_Style8
            style = workbook.Styles["Custom_Style8"];
            style.Font.IsBold = true;
            style.VerticalAlignment = TextAlignmentType.Top;
            style.Custom = "\"$\"#,##0_);[Red](\"$\"#,##0)";
            //Set style border line thickness
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

            //Custom_Style9
            style = workbook.Styles["Custom_Style9"];
            style.Number = 10;	//0.00%

            //Custom_Style10
            style = workbook.Styles["Custom_Style10"];
            style.Font.IsItalic = true;
            style.Custom = "\"$\"#,##0_);[Red](\"$\"#,##0)";
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Right;

            //Custom_Style11
            style = workbook.Styles["Custom_Style11"];
            style.Font.IsItalic = true;
            style.Number = 38;	//#,##0;[Red]-#,##0
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Right;

            //Custom_Style12
            style = workbook.Styles["Custom_Style12"];
            style.Custom = "\"$\"#,##0_);[Red](\"$\"#,##0)";
            style.HorizontalAlignment = TextAlignmentType.Right;

            //Custom_Style13
            style = workbook.Styles["Custom_Style13"];
            style.Number = 38;	//#,##0;[Red]-#,##0
            style.HorizontalAlignment = TextAlignmentType.Right;

            //Custom_Style14
            style = workbook.Styles["Custom_Style14"];
            style.Font.IsItalic = true;
            style.Number = 9;	//0%
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;

            //Custom_Style15
            style = workbook.Styles["Custom_Style15"];
            style.Number = 38;	//#,##0;[Red]-#,##0
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Right;

            //Custom_Style16
            style = workbook.Styles["Custom_Style16"];
            style.Font.IsBold = true;
            style.HorizontalAlignment = TextAlignmentType.Right;
            style.Number = 37;	//#,##0;-#,##0

            //Custom_Style17
            style = workbook.Styles["Custom_Style17"];
            style.Custom = "\"$\"#,##0_);[Red](\"$\"#,##0)";
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Right;

            //Custom_Style18
            style = workbook.Styles["Custom_Style18"];
            style.ForegroundColor = Color.Black;
            style.Pattern = BackgroundType.Solid;

            //Custom_Style19
            style = workbook.Styles["Custom_Style19"];
            style.Custom = "0.0%";

            //Custom_Style20
            style = workbook.Styles["Custom_Style20"];
            style.Number = 10;	//0.00%
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;

            //Custom_Style21
            style = workbook.Styles["Custom_Style21"];
            style.Custom = "\"$\"#,##0_);[Red](\"$\"#,##0)";
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;

            //Custom_Style22
            style = workbook.Styles["Custom_Style22"];
            style.Number = 38;	//#,##0;[Red]-#,##0

            //Custom_Style23
            style = workbook.Styles["Custom_Style23"];
            style.Custom = "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;

            //Custom_Style24
            style = workbook.Styles["Custom_Style24"];
            style.Custom = "\"$\"#,##0_);[Red](\"$\"#,##0)";
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Right;
            style.IndentLevel = 2;

            //Custom_Style25
            style = workbook.Styles["Custom_Style25"];
            style.Custom = "\"$\"#,##0_);[Red](\"$\"#,##0)";
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Right;
            style.IndentLevel = 1;

            //Custom_Style26
            style = workbook.Styles["Custom_Style26"];
            style.Number = 38;	//#,##0;[Red]-#,##0
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Right;
            style.IndentLevel = 1;

            //Custom_Style27
            style = workbook.Styles["Custom_Style27"];
            style.Number = 38;	//#,##0;[Red]-#,##0
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Right;
            style.IndentLevel = 3;

            //Custom_Style28
            style = workbook.Styles["Custom_Style28"];
            style.Number = 38;	//#,##0;[Red]-#,##0
            style.ForegroundColor = Color.Silver;
            style.Pattern = BackgroundType.Solid;
            style.HorizontalAlignment = TextAlignmentType.Right;
            style.IndentLevel = 1;

        }

        private void FillModelInputs(Workbook workbook)
        {
            //Get a worksheet named "Model Inputs"
            Worksheet sheet = workbook.Worksheets["Model Inputs"];
            
            //Get the cells in the worksheet
            Cells cells = sheet.Cells;
            
            //Set the styleflag structure
            StyleFlag styleflag = new StyleFlag();
            styleflag.All = true;
            
            //Set value and style on Range
            //Create Range from Cells C24 to F24
            Range range = cells.CreateRange("C24", "F24");
            int[] iArray = new int[] { 2800, 1200, 1600, 0 };
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style1"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style on Range
            //Create Range from Cells from C25 to F25
            range = cells.CreateRange("C25", "F25");
            iArray = new int[] { 120, 80, 40, 0 };
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style2"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
           
            //Set formular and style on range
            //Create range from cells C26 to F26
            string[] strArray = new string[] { "=+C24*C25", "=+D24*D25", "=+E24*E25", "=+F24*F25" };
            range = cells.CreateRange("C26", "F26");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style3"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set formular and style to C28
            cells["C28"].Formula = "=SUM(C26:F26)";
            cells["C28"].SetStyle(workbook.Styles["Custom_Style3"]);
            
            //Set formular and style on range
            //create range from cells C31 to F31
            strArray = new string[] { "=C23", "=D23", "=E23", "=F23" };
            range = cells.CreateRange("C31", "F31");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style4"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style on range
            //create range from cells C32 to F32
            double[] dArray = new double[] { 0.5, 0.4, 0.25, 0 };
            range = cells.CreateRange("C32", "F32");
            cells.ImportArray(dArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style5"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set formular and style on range 
            //create range from cells C33 to F33
            strArray = new string[] { "=+C26*C32", "=+D26*D32", "=+E26*E32", "=+F26*F32" };
            range = cells.CreateRange("C33", "F33");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style3"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set formular and style to C35
            cells["C35"].Formula = "=SUM(C33:F33)";
            cells["C35"].SetStyle(workbook.Styles["Custom_Style3"]);
            
            //Set value and style to C38
            cells["C38"].PutValue(0.15);
            cells["C38"].SetStyle(workbook.Styles["Custom_Style6"]);
            
            //Set value and style to C40
            cells["C40"].PutValue(5);
            cells["C40"].SetStyle(workbook.Styles["Custom_Style7"]);
            
            //Set value and style to C42
            cells["C42"].PutValue(0.30);
            cells["C42"].SetStyle(workbook.Styles["Custom_Style6"]);
            
            //Set value and style to C44
            cells["C44"].PutValue(80000);
            cells["C44"].SetStyle(workbook.Styles["Custom_Style8"]);
        }

        private void FillProfitLoss(Workbook workbook)
        {
            //Get the worksheet named "Profit and Loss"
            Worksheet sheet = workbook.Worksheets["Profit and Loss"];
            
            //Get the cells in the worksheet
            Cells cells = sheet.Cells;
            
            //Set the styleflag struct
            StyleFlag styleflag = new StyleFlag();
            styleflag.All = true;
            
            //Set value and style from 
            object[,] obj2DArray = new object[,] {{"-",0.02,0.04,0.06,0.08},
												{"-",0.02,0.04,0.06,0.08},
												{0.005,0.005,0.005,0.005,0.005}};
            //Create range from cells E8 to I10
            Range range = cells.CreateRange("E8", "I10");
            cells.ImportTwoDimensionArray(obj2DArray, range.FirstRow, range.FirstColumn);
            range.ApplyStyle(workbook.Styles["Custom_Style9"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style on range 
            //create range from cells E15 to I15
            string[] strArray = new string[] { "=+'Model Inputs'!C28", "=+E15*(1+F8)", "=+F15*(1+G8)", "=+G15*(1+H8)", "=+H15*(1+I8)" };
            range = cells.CreateRange("E15", "I15");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style on range 
            //create range from cells E16 to I16
            strArray = new string[] { "=+'Model Inputs'!C35", "=+E16*(1+F9)", "=+F16*(1+G9)", "=+G16*(1+H9)", "=+H16*(1+I9)" };
            range = cells.CreateRange("E16", "I16");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style on range 
            //create range from cells E17 to I17
            strArray = new string[] { "=E15-E16", "=F15-F16", "=G15-G16", "=H15-H16", "=I15-I16" };
            range = cells.CreateRange("E17", "I17");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            
            //Set value and Style on range 
            //Create range from cells E19 to I20
            obj2DArray = new object[,] { { 0, 0, 10000, 0, 0 }, { 1000, 0, 0, 0, 0 } };
            range = cells.CreateRange("E19", "I20");
            cells.ImportTwoDimensionArray(obj2DArray, range.FirstRow, range.FirstColumn);
            range.ApplyStyle(workbook.Styles["Custom_Style12"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style on range 
            //create range from cells E21 to I21
            strArray = new string[] { "=+E17+E19+E20", "=+F17+F19+F20", "=+G17+G19+G20", "=+H17+H19+H20", "=+I17+I19+I20" };
            range = cells.CreateRange("E21", "I21");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            
            //Set value and style for E24
            cells["E24"].PutValue(40000);
            cells["E24"].SetStyle(workbook.Styles["Custom_Style12"]);
            
            //Set value and style on range 
            //create range from cells F24 to I24
            strArray = new string[] { "=+E24*(1+F9)", "=+F24*(1+G9)", "=+G24*(1+H9)", "=+H24*(1+I9)" };
            range = cells.CreateRange("F24", "I24");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style for E25
            cells["E25"].PutValue(60000);
            cells["E25"].SetStyle(workbook.Styles["Custom_Style13"]);
            
            //Set value and style on range 
            //create range from cells F25 to I25
            strArray = new string[] { "=+E25*(1+F9)", "=+F25*(1+G9)", "=+G25*(1+H9)", "=+H25*(1+I9)" };
            range = cells.CreateRange("F25", "I25");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style from E26 to I26
            strArray = new string[] {
										"=('Balance Sheet'!D15+'Balance Sheet'!D17+'Balance Sheet'!D18)/'Model Inputs'!C40",
										"=('Balance Sheet'!E15+'Balance Sheet'!E17+'Balance Sheet'!E18)/'Model Inputs'!C40*(1+F9)",
										"=('Balance Sheet'!F15+'Balance Sheet'!F17+'Balance Sheet'!F18)/'Model Inputs'!C40*(1+G9)",
										"=('Balance Sheet'!G15+'Balance Sheet'!G17+'Balance Sheet'!G18)/'Model Inputs'!C40*(1+H9)",
										"=('Balance Sheet'!H15+'Balance Sheet'!H17+'Balance Sheet'!H18)/'Model Inputs'!C40*(1+I9)"};
            range = cells.CreateRange("E26", "I26");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style for E27
            cells["E27"].PutValue(40000);
            cells["E27"].SetStyle(workbook.Styles["Custom_Style13"]);
            
            //Set value and style from F27 to I27
            strArray = new string[] { "=+E27*(1+F9)", "=+F27*(1+G9)", "=+G27*(1+H9)", "=+H27*(1+I9)" };
            range = cells.CreateRange("F27", "I27");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style from E28 to I28
            strArray = new string[] {
										"=+'Balance Sheet'!D18*'Model Inputs'!C38",
										"=('Balance Sheet'!E18*'Model Inputs'!C38)*(1+F9)",
										"=('Balance Sheet'!F18*'Model Inputs'!C38)*(1+G9)",
										"=('Balance Sheet'!G18*'Model Inputs'!C38)*(1+H9)",
										"=('Balance Sheet'!H18*'Model Inputs'!C38)*(1+I9)"};
            range = cells.CreateRange("E28", "I28");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style for E29
            cells["E29"].PutValue(30000);
            cells["E29"].SetStyle(workbook.Styles["Custom_Style13"]);
            
            //Set value and style from F29 to I29
            strArray = new string[] { "=+E29*(1+F9)", "=+F29*(1+G9)", "=+G29*(1+H9)", "=+H29*(1+I9)" };
            range = cells.CreateRange("F29", "I29");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style for E30
            cells["E30"].PutValue(15000);
            cells["E30"].SetStyle(workbook.Styles["Custom_Style13"]);
            
            //Set value and style from F30 to I30
            strArray = new string[] { "=+E30*(1+F9)", "=+F30*(1+G9)", "=+G30*(1+H9)", "=+H30*(1+I9)" };
            range = cells.CreateRange("F30", "I30");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style for E31
            cells["E31"].PutValue(18000);
            cells["E31"].SetStyle(workbook.Styles["Custom_Style13"]);
            
            //Set value and style from F31 to I31
            strArray = new string[] { "=+E31*(1+F9)", "=+F31*(1+G9)", "=+G31*(1+H9)", "=+H31*(1+I9)" };
            range = cells.CreateRange("F31", "I31");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style for E32
            cells["E32"].PutValue(4000);
            cells["E32"].SetStyle(workbook.Styles["Custom_Style13"]);
            
            //Set value and style from F32 to I32
            strArray = new string[] { "=+E32*(1+F9)", "=+F32*(1+G9)", "=+G32*(1+H9)", "=+H32*(1+I9)" };
            range = cells.CreateRange("F32", "I32");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style from E33 to I33
            strArray = new string[] { "=SUM(E24:E32)", "=SUM(F24:F32)", "=SUM(G24:G32)", "=SUM(H24:H32)", "=SUM(I24:I32)" };
            range = cells.CreateRange("E33", "I33");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            
            //Set value and style from E35 to I35
            strArray = new string[] { "=E21-E33", "=F21-F33", "=G21-G33", "=H21-H33", "=I21-I33" };
            range = cells.CreateRange("E35", "I35");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style from E37 to I37
            strArray = new string[] {
										"=-SUM('Loan Payment Calculator'!F12:F23)","=-SUM('Loan Payment Calculator'!F24:F35)",
										"=-SUM('Loan Payment Calculator'!F36:F47)","=-SUM('Loan Payment Calculator'!F48:F59)",
										"=-SUM('Loan Payment Calculator'!F60:F71)"};
            range = cells.CreateRange("E37", "I37");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style from E39 to I39
            strArray = new string[] { "=+E35-E37", "=+F35-F37", "=+G35-G37", "=+H35-H37", "=+I35-I37" };
            range = cells.CreateRange("E39", "I39");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style from E41 to I42
            obj2DArray = new object[,] { { 0, 0, 1000, 0, 0 }, { 0, 0, 0, 0, 0 } };
            range = cells.CreateRange("E41", "I42");
            cells.ImportTwoDimensionArray(obj2DArray, range.FirstRow, range.FirstColumn);
            range.ApplyStyle(workbook.Styles["Custom_Style13"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style from E44 to I44
            strArray = new string[] { "=+E39+E41+E42", "=+F39+F41+F42", "=+G39+G41+G42", "=+H39+H41+H42", "=+I39+I41+I42" };
            range = cells.CreateRange("E44", "I44");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            
            //Set value and style for D46
            cells["D46"].Formula = "='Model Inputs'!C42";
            cells["D46"].SetStyle(workbook.Styles["Custom_Style14"]);
            
            //Set value and style from E46 to I46
            strArray = new string[] {
										"=IF(E44<0,0,D46*E44)","=IF(F44<0,0,D46*F44)",
										"=IF(G44<0,0,D46*G44)","=IF(H44<0,0,D46*H44)",
										"=IF(I44<0,0,D46*I44)"};
            range = cells.CreateRange("E46", "I46");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //Set value and style from E48 to I48
            strArray = new string[] { "=E44-E46", "=F44-F46", "=G44-G46", "=H44-H46", "=I44-I46" };
            range = cells.CreateRange("E48", "I48");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black);
        }

        private void FillBalanceSheet(Workbook workbook)
        {
            //Get the worksheet "Balance Sheet"
            Worksheet sheet = workbook.Worksheets["Balance Sheet"];
            
            //Get the cells in the worksheet
            Cells cells = sheet.Cells;
            
            //Set the styleflag stuct
            StyleFlag styleflag = new StyleFlag();
            styleflag.All = true;

            //Set value(s) and style(s) for cell(s) 

            //D7
            cells["D7"].PutValue(50000);
            cells["D7"].SetStyle(workbook.Styles["Custom_Style12"]);
            
            //E7 to I7
            string[] strArray = new string[] { "=+'Cash Flow'!D36", "=+'Cash Flow'!E36", "=+'Cash Flow'!F36", "=+'Cash Flow'!G36", "=+'Cash Flow'!H36" };
            Range range = cells.CreateRange("E7", "I7");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
           
            //D8 to D12
            int[] iArray = new int[] { 3000, 25000, 0, 0, 5000 };
            range = cells.CreateRange("D8", "D12");
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, true);
            range.ApplyStyle(workbook.Styles["Custom_Style13"], styleflag);
            
            //E8 to I8
            strArray = new string[] { "=D8", "=E8", "=F8", "=G8", "=H8" };
            range = cells.CreateRange("E8", "I8");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E9 to I9
            strArray = new string[] { "=+D9", "=+E9", "=+F9", "=+G9", "=+H9" };
            range = cells.CreateRange("E9", "I9");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E10 to I10
            strArray = new string[] { "=+D10", "=+E10", "=+F10", "=+G10", "=+H10" };
            range = cells.CreateRange("E10", "I10");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E11 to I11
            strArray = new string[] { "=+D11", "=+E11", "=+F11", "=+G11", "=+H11" };
            range = cells.CreateRange("E11", "I11");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E12 to I12
            strArray = new string[] { "=+D12", "=+E12", "=+F12", "=+G12", "=+H12" };
            range = cells.CreateRange("E12", "I12");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D13 to I13
            strArray = new string[] { "=SUM(D7:D12)", "=SUM(E7:E12)", "=SUM(F7:F12)", "=SUM(G7:G12)", "=SUM(H7:H12)", "=SUM(I7:I12)" };
            range = cells.CreateRange("D13", "I13");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            
            //D15
            cells["D15"].PutValue(100000);
            cells["D15"].SetStyle(workbook.Styles["Custom_Style12"]);
            
            //E15 to I15
            strArray = new string[] { "=+D15", "=+E15", "=+F15", "=+G15", "=+H15" };
            range = cells.CreateRange("E15", "I15");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D16 to D19
            iArray = new int[] { 100000, 0, 100000, 0 };
            range = cells.CreateRange("D16", "D19");
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, true);
            range.ApplyStyle(workbook.Styles["Custom_Style13"], styleflag);
            
            //E16 to I16
            strArray = new string[] { "=+D16", "=+E16", "=+F16", "=+G16", "=+H16" };
            range = cells.CreateRange("E16", "I16");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E17 to I17
            strArray = new string[] { "=+D17", "=+E17", "=+F17", "=+G17", "=+H17" };
            range = cells.CreateRange("E17", "I17");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E18 to I18
            strArray = new string[] { "=D18", "=E18", "=F18", "=G18", "=H18" };
            range = cells.CreateRange("E18", "I18");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E19 to I19
            strArray = new string[] { "=+D19+'Profit and Loss'!E26", "=+E19+'Profit and Loss'!F26", "=+F19+'Profit and Loss'!G26", "=+G19+'Profit and Loss'!H26", "=+H19+'Profit and Loss'!I26" };
            range = cells.CreateRange("E19", "I19");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D20 to I20
            strArray = new string[] { "=SUM(D15:D18)-D19", "=SUM(E15:E18)-E19", "=SUM(F15:F18)-F19", "=SUM(G15:G18)-G19", "=SUM(H15:H18)-H19", "=SUM(I15:I18)-I19" };
            range = cells.CreateRange("D20", "I20");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            
            //D22
            cells["D22"].PutValue(0);
            cells["D22"].SetStyle(workbook.Styles["Custom_Style12"]);
            
            //E22 to I22
            strArray = new string[] { "=+D22", "=+E22", "=+F22", "=+G22", "=+H22" };
            range = cells.CreateRange("E22", "I22");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D23 to D26
            iArray = new int[] { 0, 0, 0, 0 };
            range = cells.CreateRange("D23", "D26");
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, true);
            range.ApplyStyle(workbook.Styles["Custom_Style13"], styleflag);
            
            //E23 to I23
            strArray = new string[] { "=+D23", "=+E23", "=+F23", "=+G23", "=+H23" };
            range = cells.CreateRange("E23", "I23");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
           
            //E24 to I24
            strArray = new string[] { "=+D24", "=+E24", "=+F24", "=+G24", "=+H24" };
            range = cells.CreateRange("E24", "I24");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E25 to I25
            strArray = new string[] { "=+D25", "=+E25", "=+F25", "=+G25", "=+H25" };
            range = cells.CreateRange("E25", "I25");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E26 to I26
            strArray = new string[] { "=+D26", "=+E26", "=+F26", "=+G26", "=+H26" };
            range = cells.CreateRange("E26", "I26");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D27 to I27
            strArray = new string[] { "=SUM(D22:D26)+D20+D13", "=SUM(E22:E26)+E20+E13", "=SUM(F22:F26)+F20+F13", "=SUM(G22:G26)+G20+G13", "=SUM(H22:H26)+H20+H13", "=SUM(I22:I26)+I20+I13" };
            range = cells.CreateRange("D27", "I27");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black);
            
            //D30
            cells["D30"].PutValue(0);
            cells["D30"].SetStyle(workbook.Styles["Custom_Style12"]);
           
            //E30 to I30
            strArray = new string[] { "=+D30", "=+E30", "=+F30", "=+G30", "=+H30" };
            range = cells.CreateRange("E30", "I30");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D31 to D34
            iArray = new int[] { 0, 0, 0, 100 };
            range = cells.CreateRange("D31", "D34");
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, true);
            range.ApplyStyle(workbook.Styles["Custom_Style13"], styleflag);
            
            //E31 to I31
            strArray = new string[] { "=+D31", "=+E31", "=+F31", "=+G31", "=+H31" };
            range = cells.CreateRange("E31", "I31");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E32 to I32
            strArray = new string[] { "=+D32", "=+E32", "=+F32", "=+G32", "=+H32" };
            range = cells.CreateRange("E32", "I32");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
           
            //E33 to I33
            strArray = new string[] { "=+D33", "=+E33", "=+F33", "=+G33", "=+H33" };
            range = cells.CreateRange("E33", "I33");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E34 to I34
            strArray = new string[] { "=+D34", "=+E34", "=+F34", "=+G34", "=+H34" };
            range = cells.CreateRange("E34", "I34");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D35 to I35
            strArray = new string[] { "=SUM(D30:D34)", "=SUM(E30:E34)", "=SUM(F30:F34)", "=SUM(G30:G34)", "=SUM(H30:H34)", "=SUM(I30:I34)" };
            range = cells.CreateRange("D35", "I35");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            
            //D37 to I37
            strArray = new string[] { "=+'Model Inputs'!C44", "='Loan Payment Calculator'!D24", "='Loan Payment Calculator'!D36", "='Loan Payment Calculator'!D48", "='Loan Payment Calculator'!D60", "='Loan Payment Calculator'!D72" };
            range = cells.CreateRange("D37", "I37");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D38 to I38
            iArray = new int[] { 100000, 200000, 150000, 175000, 225000, 150000 };
            range = cells.CreateRange("D38", "I38");
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style12"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D39 to I39
            strArray = new string[] { "=D35+D37+D38", "=E35+E37+E38", "=F35+F37+F38", "=G35+G37+G38", "=H35+H37+H38", "=I35+I37+I38" };
            range = cells.CreateRange("D39", "I39");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            
            //D41
            cells["D41"].PutValue(0);
            cells["D41"].SetStyle(workbook.Styles["Custom_Style16"]);
            
            //E41 to I41
            strArray = new string[] { "=+D41", "=+E41", "=+F41", "=+G41", "=+H41" };
            range = cells.CreateRange("E41", "I41");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D43 to I43
            strArray = new string[] { "=D35+D37+D41", "=E35+E37+E41", "=F35+F37+F41", "=G35+G37+G41", "=H35+H37+H41", "=I35+I37+I41" };
            range = cells.CreateRange("D43", "I43");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black);
            
            //D46
            cells["D46"].PutValue(50000);
            cells["D46"].SetStyle(workbook.Styles["Custom_Style12"]);
            
            //E46 to I46
            strArray = new string[] { "=D46", "=E46", "=F46", "=G46", "=H46" };
            range = cells.CreateRange("E46", "I46");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style16"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D47 to D49
            iArray = new int[] { 250000, 0, 0 };
            range = cells.CreateRange("D47", "D49");
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, true);
            range.ApplyStyle(workbook.Styles["Custom_Style13"], styleflag);
            
            //E47 to I47
            strArray = new string[] { "=+D47", "=+E47", "=+F47", "=+G47", "=+H47" };
            range = cells.CreateRange("E47", "I47");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E48 to I48
            strArray = new string[] { "=+D48", "=+E48", "=+F48", "=+G48", "=+H48" };
            range = cells.CreateRange("E48", "I48");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E49 to I49
            strArray = new string[] { "=+'Profit and Loss'!E48", "=+E49+'Profit and Loss'!F48", "=+F49+'Profit and Loss'!G48", "=+G49+'Profit and Loss'!H48", "=+H49+'Profit and Loss'!I48" };
            range = cells.CreateRange("E49", "I49");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D50 to I50
            strArray = new string[] { "=SUM(D46:D49)", "=SUM(E46:E49)", "=SUM(F46:F49)", "=SUM(G46:G49)", "=SUM(H46:H49)", "=SUM(I46:I49)" };
            range = cells.CreateRange("D50", "I50");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black);
           
            //D52 to I52
            strArray = new string[] { "=D50+D43", "=E50+E43", "=F50+F43", "=G50+G43", "=H50+H43", "=I50+I43" };
            range = cells.CreateRange("D52", "I52");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black);
        }

        private void FillCashFlow(Workbook workbook)
        {
            //Get the worksheet named  "Cash Flow"
            Worksheet sheet = workbook.Worksheets["Cash Flow"];
            
            //Get the cells collection in the sheet
            //Set the styleflag stuct
            Cells cells = sheet.Cells;
            StyleFlag styleflag = new StyleFlag();
            styleflag.All = true;

            //Set value(s) and style(s) for cell(s)

            //D8 to I8
            string[] strArray = new string[] { "=+'Profit and Loss'!E48", "=+'Profit and Loss'!F48", "=+'Profit and Loss'!G48", "=+'Profit and Loss'!H48", "=+'Profit and Loss'!I48", "=SUM(D8:H8)" };
            Range range = cells.CreateRange("D8", "I8");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D9 to I9
            strArray = new string[] { "='Profit and Loss'!E26", "='Profit and Loss'!F26", "='Profit and Loss'!G26", "='Profit and Loss'!H26", "='Profit and Loss'!I26", "=SUM(D9:H9)" };
            range = cells.CreateRange("D9", "I9");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
           
            //D10 to I10
            strArray = new string[] { "=+'Balance Sheet'!D8-'Balance Sheet'!E8", "=+'Balance Sheet'!D8-'Balance Sheet'!F8", "=+'Balance Sheet'!D8-'Balance Sheet'!G8", "=+'Balance Sheet'!D8-'Balance Sheet'!H8", "=+'Balance Sheet'!D8-'Balance Sheet'!I8", "=SUM(D10:H10)" };
            range = cells.CreateRange("D10", "I10");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D11 to I11
            strArray = new string[] { "=+'Balance Sheet'!D9-'Balance Sheet'!E9", "=+'Balance Sheet'!D9-'Balance Sheet'!F9", "=+'Balance Sheet'!D9-'Balance Sheet'!G9", "=+'Balance Sheet'!D9-'Balance Sheet'!H9", "=+'Balance Sheet'!D9-'Balance Sheet'!I9", "=SUM(D11:H11)" };
            range = cells.CreateRange("D11", "I11");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D12 to I12
            strArray = new string[] { "=+'Balance Sheet'!D30-'Balance Sheet'!E30", "=+'Balance Sheet'!D30-'Balance Sheet'!F30", "=+'Balance Sheet'!D30-'Balance Sheet'!G30", "=+'Balance Sheet'!D30-'Balance Sheet'!H30", "=+'Balance Sheet'!D30-'Balance Sheet'!I30", "=SUM(D12:H12)" };
            range = cells.CreateRange("D12", "I12");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
           
            //D13 to D15
            int[] iArray = new int[] { 0, 0, 0 };
            range = cells.CreateRange("D13", "D15");
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, true);
            range.ApplyStyle(workbook.Styles["Custom_Style13"], styleflag);
           
            //E13 to I13
            strArray = new string[] { "=D13", "=E13", "=F13", "=G13", "=SUM(D13:H13)" };
            range = cells.CreateRange("E13", "I13");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E14 to I14
            strArray = new string[] { "=D14", "=E14", "=F14", "=G14", "=SUM(D14:H14)" };
            range = cells.CreateRange("E14", "I14");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E15 to I15
            strArray = new string[] { "=D15", "=E15", "=F15", "=G15", "=SUM(D15:H15)" };
            range = cells.CreateRange("E15", "I15");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D16 to I16
            strArray = new string[] { "=SUM(D8:D15)", "=SUM(E8:E15)", "=SUM(F8:F15)", "=SUM(G8:G15)", "=SUM(H8:H15)", "=SUM(D16:H16)" };
            range = cells.CreateRange("D16", "I16");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            range = cells.CreateRange("I8", "I16");
            range.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            
            //D19 to I19
            strArray = new string[] { "=-1 * MAX(0,SUM('Balance Sheet'!E15:E18)-SUM('Balance Sheet'!D15:D18))", "=-1 * MAX(0,SUM('Balance Sheet'!F15:F18)-SUM('Balance Sheet'!E15:E18))", "=-1 * MAX(0,SUM('Balance Sheet'!G15:G18)-SUM('Balance Sheet'!F15:F18))", "=-1 * MAX(0,SUM('Balance Sheet'!H15:H18)-SUM('Balance Sheet'!G15:G18))", "=-1 * MAX(0,SUM('Balance Sheet'!I15:I18)-SUM('Balance Sheet'!H15:H18))", "=SUM(D19:H19)" };
            range = cells.CreateRange("D19", "I19");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D20 to H20
            iArray = new int[] { 0, 0, 0, 0, 0 };
            range = cells.CreateRange("D20", "H20");
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style13"], styleflag);
            
            //I20 to I22
            strArray = new string[] { "=SUM(D20:H20)", "=SUM(D21:H21)", "=SUM(D22:H22)" };
            range = cells.CreateRange("I20", "I22");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, true);
            range.ApplyStyle(workbook.Styles["Custom_Style11"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D21 to H21
            strArray = new string[] { "=(-1*MIN(0,SUM('Balance Sheet'!E15:E18)-SUM('Balance Sheet'!D15:D18))-'Profit and Loss'!E41)", "=(-1*MIN(0,SUM('Balance Sheet'!F15:F18)-SUM('Balance Sheet'!E15:E18))-'Profit and Loss'!F41)", "=(-1*MIN(0,SUM('Balance Sheet'!G15:G18)-SUM('Balance Sheet'!F15:F18))-'Profit and Loss'!G41)", "=(-1*MIN(0,SUM('Balance Sheet'!H15:H18)-SUM('Balance Sheet'!G15:G18))-'Profit and Loss'!H41)", "=(-1*MIN(0,SUM('Balance Sheet'!I15:I18)-SUM('Balance Sheet'!H15:H18))-'Profit and Loss'!I41)" };
            range = cells.CreateRange("D21", "H21");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            
            //D22 to H22
            iArray = new int[] { 0, 0, 0, 0, 0 };
            range = cells.CreateRange("D22", "H22");
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style13"], styleflag);
            
            //D23 to I23
            strArray = new string[] { "=SUM(D19:D22)", "=SUM(E19:E22)", "=SUM(F19:F22)", "=SUM(G19:G22)", "=SUM(H19:H22)", "=SUM(D23:H23)" };
            range = cells.CreateRange("D23", "I23");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

            range = cells.CreateRange("I19", "I23");
            range.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);

            //D26 to I26
            strArray = new string[] { "= 'Balance Sheet'!E37+'Balance Sheet'!E38-('Balance Sheet'!D37+'Balance Sheet'!D38)", "= 'Balance Sheet'!F37+'Balance Sheet'!F38-('Balance Sheet'!E37+'Balance Sheet'!E38)", "= 'Balance Sheet'!G37+'Balance Sheet'!G38-('Balance Sheet'!F37+'Balance Sheet'!F38)", "= 'Balance Sheet'!H37+'Balance Sheet'!H38-('Balance Sheet'!G37+'Balance Sheet'!G38)", "= 'Balance Sheet'!I37+'Balance Sheet'!I38-('Balance Sheet'!H37+'Balance Sheet'!H38)", "=SUM(D26:H26)" };
            range = cells.CreateRange("D26", "I26");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D27 to D30
            iArray = new int[] { 0, 0, 0, 0 };
            range = cells.CreateRange("D27", "D30");
            cells.ImportArray(iArray, range.FirstRow, range.FirstColumn, true);
            range.ApplyStyle(workbook.Styles["Custom_Style13"], styleflag);
            
            //E27 to I27
            strArray = new string[] { "=D27", "=E27", "=F27", "=G27", "=SUM(D27:H27)" };
            range = cells.CreateRange("E27", "I27");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E28 to I28
            strArray = new string[] { "=D28", "=E28", "=F28", "=G28", "=SUM(D28:H28)" };
            range = cells.CreateRange("E28", "I28");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //E29 to I29
            strArray = new string[] { "=D29", "=E29", "=F29", "=G29", "=SUM(D29:H29)" };
            range = cells.CreateRange("E29", "I29");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
           
            //E30 to I30
            strArray = new string[] { "=D30", "=E30", "=F30", "=G30", "=SUM(D30:H30)" };
            range = cells.CreateRange("E30", "I30");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style15"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D31 to I31
            strArray = new string[] { "=SUM(D26:D30)", "=SUM(E26:E30)", "=SUM(F26:F30)", "=SUM(G26:G30)", "=SUM(H26:H30)", "=SUM(D31:H31)" };
            range = cells.CreateRange("D31", "I31");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

            range = cells.CreateRange("I26", "I31");
            range.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            
            //D33 to I33
            strArray = new string[] { "=D16+D23+D31", "=E16+E23+E31", "=F16+F23+F31", "=G16+G23+G31", "=H16+H23+H31", "=SUM(D33:H33)" };
            range = cells.CreateRange("D33", "I33");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            
            //D35 to H35
            strArray = new string[] { "=+'Balance Sheet'!D7", "=D35", "=E35", "=F35", "=G35" };
            range = cells.CreateRange("D35", "H35");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            
            //D36 to H36
            strArray = new string[] { "=D35+D33", "=E35+E33", "=F35+F33", "=G35+G33", "=H35+H33" };
            range = cells.CreateRange("D36", "H36");
            cells.ImportFormulaArray(strArray, range.FirstRow, range.FirstColumn, false);
            range.ApplyStyle(workbook.Styles["Custom_Style10"], styleflag);
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);

            range = cells.CreateRange("I33", "I36");
            range.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            
            //I34 to I36
            range = cells.CreateRange("I34", "I36");
            range.ApplyStyle(workbook.Styles["Custom_Style18"], styleflag);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);
        }

        private void FillLoanPaymentCalculator(Workbook workbook)
        {
            //Get the worksheet named "Loan Payment Calculator"
            Worksheet sheet = workbook.Worksheets["Loan Payment Calculator"];
            
            //Get the cells in the sheet
            Cells cells = sheet.Cells;
            
            //Input values, set formulas and apply styles to cells 
            cells["C5"].PutValue(0.05);
            cells["C5"].SetStyle(workbook.Styles["Custom_Style19"]);
            cells["C6"].Formula = "=(1+C5)^(1/12)-1";
            cells["C6"].SetStyle(workbook.Styles["Custom_Style20"]);
            cells["C7"].Formula = "=+'Model Inputs'!C44";
            cells["C7"].SetStyle(workbook.Styles["Custom_Style21"]);
            cells["C8"].PutValue(60);
            cells["C8"].SetStyle(workbook.Styles["Custom_Style22"]);
            cells["C9"].Formula = "=PMT(C6,C8,C7)";
            cells["C9"].SetStyle(workbook.Styles["Custom_Style23"]);
            Range range = cells.CreateRange("C5", "C9");
            range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

            //Set headers
            cells["D12"].Formula = "=C7";
            cells["D12"].SetStyle(workbook.Styles["Custom_Style24"]);
            cells["E12"].Formula = "=G12-F12";
            cells["E12"].SetStyle(workbook.Styles["Custom_Style23"]);
            cells["F12"].Formula = "=-C6*D12";
            cells["F12"].SetStyle(workbook.Styles["Custom_Style25"]);
            cells["G12"].Formula = "=IF(C12>C8, 0, C9)";
            cells["G12"].SetStyle(workbook.Styles["Custom_Style25"]);
            range = cells.CreateRange("D12", "G12");
            range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

            int k = 12;
            for (int i = 2; i <= 360; i++, k++)
            {
                int j = k + 1;
                //Set style and formula for column D
                cells["D" + j.ToString()].Formula = "=D" + k.ToString() + "+E" + k.ToString();
                cells["D" + j.ToString()].SetStyle(workbook.Styles["Custom_Style26"]);
               
                //Set style and formula for column E
                cells["E" + j.ToString()].Formula = "=G" + j.ToString() + "-F" + j.ToString();
                cells["E" + j.ToString()].SetStyle(workbook.Styles["Custom_Style27"]);
                
                //Set style and formula for column F
                cells["F" + j.ToString()].Formula = "=-C6*D" + j.ToString();
                cells["F" + j.ToString()].SetStyle(workbook.Styles["Custom_Style26"]);
               
                //Set style and formula for column G
                cells["G" + j.ToString()].Formula = "=IF(C" + j.ToString() + ">C8, 0, C9)";
                cells["G" + j.ToString()].SetStyle(workbook.Styles["Custom_Style25"]);
               
                //Set borders
                range = cells.CreateRange("D" + j.ToString(), "G" + j.ToString());
                range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            }
        }
    }
}


