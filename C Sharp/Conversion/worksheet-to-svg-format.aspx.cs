using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;


namespace Aspose.Cells.Demos.Conversion
{
    public partial class Worksheet2Svg : System.Web.UI.Page
    {

        protected void btnExecute_Click(object sender, EventArgs e)
        {
            CreateStaticReport();
        }

        public void CreateStaticReport()
        {

            //Open template
            string path = System.Web.HttpContext.Current.Server.MapPath("~");
            path = path.Substring(0, path.LastIndexOf("\\"));

            string outpath = path;

            path += @"\designer\ProductList.xls";
            outpath += @"\designer\Output\";

            //Lnks will be used to get the output files for later use
            ArrayList lnks = new ArrayList();


            //Create a workbook object from the template file
            Workbook book = new Workbook(path);

            //Convert each worksheet into svg format in a single page.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            imgOptions.SaveFormat = SaveFormat.SVG;
            imgOptions.OnePagePerSheet = true;

            
            //Convert each worksheet into svg format
            foreach (Worksheet worksheet in book.Worksheets)
            {
                SheetRender sr = new SheetRender(worksheet, imgOptions);
               
                for (int i = 0; i < sr.PageCount; i++)
                {
                    
                    string svgFileName = "ProductList" + worksheet.Index + i + ".svg";
                    lnks.Add(svgFileName);

                    //Output the worksheet into Svg format
                    sr.ToImage(i, outpath + svgFileName);                    
                }
            }


            //Show links to all output svg files
            Literal ltr=new Literal();
            ltr.Text="<ul>";

            outPanel.Controls.Add(ltr);

            foreach (string lnk in lnks)
            {
                ltr = new Literal();
                ltr.Text = "<li>";
                outPanel.Controls.Add(ltr);

                HyperLink hyp = new HyperLink();
                hyp.ID = lnk;
                hyp.Text = lnk; ;
                hyp.NavigateUrl = "~/designer/Output/" + lnk;
                outPanel.Controls.Add(hyp);

                ltr = new Literal();
                ltr.Text = "</li>";
                outPanel.Controls.Add(ltr);
            }

            ltr = new Literal();
            ltr.Text = "</ul>";
            outPanel.Controls.Add(ltr);

        }
    }
}
