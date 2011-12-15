using System;
using System.Data;
using System.IO;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Aspose.Cells;
using Aspose.Cells.Rendering;

public partial class Sheet2ImageWithPrintArea : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {


    }
    protected void btnExecute_Click(object sender, EventArgs e)
    {
        CreateStaticReport();
    }

    public static void CreateStaticReport()
    {
        //Instantiating an Workbook object
        Workbook workbook = new Workbook();

        //Obtaining the reference of the newly added worksheet
        Worksheet sheet = workbook.Worksheets[0];

        Cells cells = sheet.Cells;

        //Setting the value to the cells
        Aspose.Cells.Cell cell = cells["A1"];
        cell.PutValue("Employee");
        cell = cells["B1"];
        cell.PutValue("Quarter");
        cell = cells["C1"];
        cell.PutValue("Product");
        cell = cells["D1"];
        cell.PutValue("Continent");
        cell = cells["E1"];
        cell.PutValue("Country");
        cell = cells["F1"];
        cell.PutValue("Sale");

        cell = cells["A2"];
        cell.PutValue("David");
        cell = cells["A3"];
        cell.PutValue("David");
        cell = cells["A4"];
        cell.PutValue("David");
        cell = cells["A5"];
        cell.PutValue("David");
        cell = cells["A6"];
        cell.PutValue("James");
        cell = cells["A7"];
        cell.PutValue("James");
        cell = cells["A8"];
        cell.PutValue("James");
        cell = cells["A9"];
        cell.PutValue("James");
        cell = cells["A10"];
        cell.PutValue("James");
        cell = cells["A11"];
        cell.PutValue("Miya");
        cell = cells["A12"];
        cell.PutValue("Miya");
        cell = cells["A13"];
        cell.PutValue("Miya");
        cell = cells["A14"];
        cell.PutValue("Miya");
        cell = cells["A15"];
        cell.PutValue("Miya");
        cell = cells["A16"];
        cell.PutValue("Miya");
        cell = cells["A17"];
        cell.PutValue("Miya");
        cell = cells["A18"];
        cell.PutValue("Elvis");
        cell = cells["A19"];
        cell.PutValue("Elvis");
        cell = cells["A20"];
        cell.PutValue("Elvis");
        cell = cells["A21"];
        cell.PutValue("Elvis");
        cell = cells["A22"];
        cell.PutValue("Elvis");
        cell = cells["A23"];
        cell.PutValue("Elvis");
        cell = cells["A24"];
        cell.PutValue("Elvis");
        cell = cells["A25"];
        cell.PutValue("Jean");
        cell = cells["A26"];
        cell.PutValue("Jean");
        cell = cells["A27"];
        cell.PutValue("Jean");
        cell = cells["A28"];
        cell.PutValue("Ada");
        cell = cells["A29"];
        cell.PutValue("Ada");
        cell = cells["A30"];
        cell.PutValue("Ada");

        cell = cells["B2"];
        cell.PutValue("1");
        cell = cells["B3"];
        cell.PutValue("2");
        cell = cells["B4"];
        cell.PutValue("3");
        cell = cells["B5"];
        cell.PutValue("4");
        cell = cells["B6"];
        cell.PutValue("1");
        cell = cells["B7"];
        cell.PutValue("2");
        cell = cells["B8"];
        cell.PutValue("3");
        cell = cells["B9"];
        cell.PutValue("4");
        cell = cells["B10"];
        cell.PutValue("4");
        cell = cells["B11"];
        cell.PutValue("1");
        cell = cells["B12"];
        cell.PutValue("1");
        cell = cells["B13"];
        cell.PutValue("2");
        cell = cells["B14"];
        cell.PutValue("2");
        cell = cells["B15"];
        cell.PutValue("3");
        cell = cells["B16"];
        cell.PutValue("4");
        cell = cells["B17"];
        cell.PutValue("4");
        cell = cells["B18"];
        cell.PutValue("1");
        cell = cells["B19"];
        cell.PutValue("1");
        cell = cells["B20"];
        cell.PutValue("2");
        cell = cells["B21"];
        cell.PutValue("3");
        cell = cells["B22"];
        cell.PutValue("3");
        cell = cells["B23"];
        cell.PutValue("4");
        cell = cells["B24"];
        cell.PutValue("4");
        cell = cells["B25"];
        cell.PutValue("1");
        cell = cells["B26"];
        cell.PutValue("2");
        cell = cells["B27"];
        cell.PutValue("3");
        cell = cells["B28"];
        cell.PutValue("1");
        cell = cells["B29"];
        cell.PutValue("2");
        cell = cells["B30"];
        cell.PutValue("3");

        cell = cells["C2"];
        cell.PutValue("Maxilaku");
        cell = cells["C3"];
        cell.PutValue("Maxilaku");
        cell = cells["C4"];
        cell.PutValue("Chai");
        cell = cells["C5"];
        cell.PutValue("Maxilaku");
        cell = cells["C6"];
        cell.PutValue("Chang");
        cell = cells["C7"];
        cell.PutValue("Chang");
        cell = cells["C8"];
        cell.PutValue("Chang");
        cell = cells["C9"];
        cell.PutValue("Chang");
        cell = cells["C10"];
        cell.PutValue("Chang");
        cell = cells["C11"];
        cell.PutValue("Geitost");
        cell = cells["C12"];
        cell.PutValue("Chai");
        cell = cells["C13"];
        cell.PutValue("Geitost");
        cell = cells["C14"];
        cell.PutValue("Geitost");
        cell = cells["C15"];
        cell.PutValue("Maxilaku");
        cell = cells["C16"];
        cell.PutValue("Geitost");
        cell = cells["C17"];
        cell.PutValue("Geitost");
        cell = cells["C18"];
        cell.PutValue("Ikuru");
        cell = cells["C19"];
        cell.PutValue("Ikuru");
        cell = cells["C20"];
        cell.PutValue("Ikuru");
        cell = cells["C21"];
        cell.PutValue("Ikuru");
        cell = cells["C22"];
        cell.PutValue("Ipoh Coffee");
        cell = cells["C23"];
        cell.PutValue("Ipoh Coffee");
        cell = cells["C24"];
        cell.PutValue("Ipoh Coffee");
        cell = cells["C25"];
        cell.PutValue("Chocolade");
        cell = cells["C26"];
        cell.PutValue("Chocolade");
        cell = cells["C27"];
        cell.PutValue("Chocolade");
        cell = cells["C28"];
        cell.PutValue("Chocolade");
        cell = cells["C29"];
        cell.PutValue("Chocolade");
        cell = cells["C30"];
        cell.PutValue("Chocolade");

        cell = cells["D2"];
        cell.PutValue("Asia");
        cell = cells["D3"];
        cell.PutValue("Asia");
        cell = cells["D4"];
        cell.PutValue("Asia");
        cell = cells["D5"];
        cell.PutValue("Asia");
        cell = cells["D6"];
        cell.PutValue("Europe");
        cell = cells["D7"];
        cell.PutValue("Europe");
        cell = cells["D8"];
        cell.PutValue("Europe");
        cell = cells["D9"];
        cell.PutValue("Europe");
        cell = cells["D10"];
        cell.PutValue("Europe");
        cell = cells["D11"];
        cell.PutValue("America");
        cell = cells["D12"];
        cell.PutValue("America");
        cell = cells["D13"];
        cell.PutValue("America");
        cell = cells["D14"];
        cell.PutValue("America");
        cell = cells["D15"];
        cell.PutValue("America");
        cell = cells["D16"];
        cell.PutValue("America");
        cell = cells["D17"];
        cell.PutValue("America");
        cell = cells["D18"];
        cell.PutValue("Europe");
        cell = cells["D19"];
        cell.PutValue("Europe");
        cell = cells["D20"];
        cell.PutValue("Europe");
        cell = cells["D21"];
        cell.PutValue("Oceania");
        cell = cells["D22"];
        cell.PutValue("Oceania");
        cell = cells["D23"];
        cell.PutValue("Oceania");
        cell = cells["D24"];
        cell.PutValue("Oceania");
        cell = cells["D25"];
        cell.PutValue("Africa");
        cell = cells["D26"];
        cell.PutValue("Africa");
        cell = cells["D27"];
        cell.PutValue("Africa");
        cell = cells["D28"];
        cell.PutValue("Africa");
        cell = cells["D29"];
        cell.PutValue("Africa");
        cell = cells["D30"];
        cell.PutValue("Africa");

        cell = cells["E2"];
        cell.PutValue("China");
        cell = cells["E3"];
        cell.PutValue("India");
        cell = cells["E4"];
        cell.PutValue("Korea");
        cell = cells["E5"];
        cell.PutValue("India");
        cell = cells["E6"];
        cell.PutValue("France");
        cell = cells["E7"];
        cell.PutValue("France");
        cell = cells["E8"];
        cell.PutValue("Germany");
        cell = cells["E9"];
        cell.PutValue("Italy");
        cell = cells["E10"];
        cell.PutValue("France");
        cell = cells["E11"];
        cell.PutValue("U.S.");
        cell = cells["E12"];
        cell.PutValue("U.S.");
        cell = cells["E13"];
        cell.PutValue("Brazil");
        cell = cells["E14"];
        cell.PutValue("U.S.");
        cell = cells["E15"];
        cell.PutValue("U.S.");
        cell = cells["E16"];
        cell.PutValue("Canada");
        cell = cells["E17"];
        cell.PutValue("U.S.");
        cell = cells["E18"];
        cell.PutValue("Italy");
        cell = cells["E19"];
        cell.PutValue("France");
        cell = cells["E20"];
        cell.PutValue("Italy");
        cell = cells["E21"];
        cell.PutValue("New Zealand");
        cell = cells["E22"];
        cell.PutValue("Australia");
        cell = cells["E23"];
        cell.PutValue("Australia");
        cell = cells["E24"];
        cell.PutValue("New Zealand");
        cell = cells["E25"];
        cell.PutValue("S.Africa");
        cell = cells["E26"];
        cell.PutValue("S.Africa");
        cell = cells["E27"];
        cell.PutValue("S.Africa");
        cell = cells["E28"];
        cell.PutValue("Egypt");
        cell = cells["E29"];
        cell.PutValue("Egypt");
        cell = cells["E30"];
        cell.PutValue("Egypt");

        cell = cells["F2"];
        cell.PutValue(2000);
        cell = cells["F3"];
        cell.PutValue(500);
        cell = cells["F4"];
        cell.PutValue(1200);
        cell = cells["F5"];
        cell.PutValue(1500);
        cell = cells["F6"];
        cell.PutValue(500);
        cell = cells["F7"];
        cell.PutValue(1500);
        cell = cells["F8"];
        cell.PutValue(800);
        cell = cells["F9"];
        cell.PutValue(900);
        cell = cells["F10"];
        cell.PutValue(500);
        cell = cells["F11"];
        cell.PutValue(1600);
        cell = cells["F12"];
        cell.PutValue(600);
        cell = cells["F13"];
        cell.PutValue(2000);
        cell = cells["F14"];
        cell.PutValue(500);
        cell = cells["F15"];
        cell.PutValue(900);
        cell = cells["F16"];
        cell.PutValue(700);
        cell = cells["F17"];
        cell.PutValue(1400);
        cell = cells["F18"];
        cell.PutValue(1350);
        cell = cells["F19"];
        cell.PutValue(300);
        cell = cells["F20"];
        cell.PutValue(500);
        cell = cells["F21"];
        cell.PutValue(1000);
        cell = cells["F22"];
        cell.PutValue(1500);
        cell = cells["F23"];
        cell.PutValue(1500);
        cell = cells["F24"];
        cell.PutValue(1600);
        cell = cells["F25"];
        cell.PutValue(1000);
        cell = cells["F26"];
        cell.PutValue(1200);
        cell = cells["F27"];
        cell.PutValue(1300);
        cell = cells["F28"];
        cell.PutValue(1500);
        cell = cells["F29"];
        cell.PutValue(1400);
        cell = cells["F30"];
        cell.PutValue(1000);

        //Set Page Orientation
        sheet.PageSetup.Orientation = PageOrientationType.Portrait;

        //Set Paper Size
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

        //Show Headings
        sheet.PageSetup.PrintHeadings = true;

        //Set Print Area
        sheet.PageSetup.PrintArea = "A1:C30,D1:F30";

        //Create a memory stream object.
        MemoryStream memorystream = new MemoryStream();

        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Tiff;
        
        SheetRender sheetRender = new SheetRender(sheet, imgOptions);

        //Convert worksheet to image.
        sheetRender.ToTiff(memorystream);

        memorystream.Seek(0, SeekOrigin.Begin);  

        //Set Response object to stream the image file.
        byte[] data = memorystream.ToArray();
        HttpContext.Current.Response.Clear();
        HttpContext.Current.Response.ContentType = "image/tiff";
        HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=SheetImage.tiff");
        HttpContext.Current.Response.OutputStream.Write(data, 0, data.Length);

        //End response to avoid unneeded html after xls
        HttpContext.Current.Response.End();

    }


}
