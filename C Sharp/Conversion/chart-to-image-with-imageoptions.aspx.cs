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
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;


public partial class Chart2ImageWithOptions : System.Web.UI.Page
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
        //Create a new Workbook.
        Workbook workbook = new Workbook();
       
        //Get the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];
        
        //Set the name of worksheet
        sheet.Name = "ChartSheet";
        
        //Get the cells collection in the sheet.
        Cells cells = workbook.Worksheets[0].Cells;
        
        //Put some values into different cells of the sheet.
        cells["A1"].PutValue("Region");
        cells["A2"].PutValue("France");
        cells["A3"].PutValue("Germany");
        cells["A4"].PutValue("England");
        cells["A5"].PutValue("Sweden");
        cells["A6"].PutValue("Italy");
        cells["A7"].PutValue("Spain");
        cells["A8"].PutValue("Portugal");
        cells["B1"].PutValue("Sale");
        cells["B2"].PutValue(70000);
        cells["B3"].PutValue(55000);
        cells["B4"].PutValue(30000);
        cells["B5"].PutValue(40000);
        cells["B6"].PutValue(35000);
        cells["B7"].PutValue(32000);
        cells["B8"].PutValue(10000);

        //Create chart
        int chartIndex = 0;
        chartIndex = sheet.Charts.Add(ChartType.Pie, 2, 4, 31, 15);
        Chart chart = sheet.Charts[chartIndex];

        //Set properties of chart title
        chart.Title.Text = "Sales By Region";
        chart.Title.TextFont.Color = System.Drawing.Color.Blue;
        chart.Title.TextFont.IsBold = true;
        chart.Title.TextFont.Size = 12;

        //Set properties of nseries
        chart.NSeries.Add("B2:B8", true);
        chart.NSeries.CategoryData = "A2:A8";
        chart.NSeries.IsColorVaried = false;

        for (int i = 0; i < chart.NSeries.Count; i++)
        {
            //Set the DataLabels in the chart
            Aspose.Cells.Charts.DataLabels dataLabels = chart.NSeries[i].DataLabels;
            dataLabels.Position = LabelPositionType.OutsideEnd;
		    dataLabels.ShowCategoryName = false;
            dataLabels.ShowValue = false;
            dataLabels.ShowPercentage = true;
            dataLabels.ShowLegendKey = false;

        }

        //Set the Legend.
        Legend legend = chart.Legend;
        legend.Position = LegendPositionType.Left;

        //Apply different Image and Print options 
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.HorizontalResolution = 300;
        options.VerticalResolution = 300;
        options.TiffCompression = TiffCompression.CompressionLZW;
        options.IsCellAutoFit = false;
        options.ImageFormat = System.Drawing.Imaging.ImageFormat.Tiff;
        options.PrintingPage = PrintingPageType.Default;

        //Create a memory stream object.
        MemoryStream ms = new MemoryStream();
        
        //Conver the chart to image file.
        chart.ToImage(ms, options);
       
        //Set Response object to stream the image file.
        byte[] data = ms.ToArray();
        HttpContext.Current.Response.Clear();
        HttpContext.Current.Response.ContentType = "image/tiff";
        HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=ChartPic.tiff");
        HttpContext.Current.Response.OutputStream.Write(data, 0, data.Length);
        
        //End response to avoid unneeded html after xls
        HttpContext.Current.Response.End();   
    }


}
