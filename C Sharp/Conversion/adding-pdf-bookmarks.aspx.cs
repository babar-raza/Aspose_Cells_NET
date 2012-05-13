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

public partial class AddingPdfBookmarks : System.Web.UI.Page
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
        //Instantiate a new workbook
        Workbook workbook = new Workbook();

        //Get the cells in the first(default) worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        //Get the A1 cell
        Aspose.Cells.Cell p = cells["A1"];

        //Enter a value
        p.PutValue("Preface");

        //Get the A10 cell
        Aspose.Cells.Cell A = cells["A10"];

        //Enter a value.
        A.PutValue("page1");

        //Get the H15 cell
        Aspose.Cells.Cell D = cells["H15"];

        //Enter a value
        D.PutValue("page1(H15)");

        //Add a new worksheet to the workbook
        workbook.Worksheets.Add();

        //Get the cells in the second sheet
        cells = workbook.Worksheets[1].Cells;

        //Get the B10 cell in the second sheet
        Aspose.Cells.Cell B = cells["B10"];

        //Enter a value
        B.PutValue("page2");

        //Add a new worksheet to the workbook
        workbook.Worksheets.Add();

        //Get the cells in the third sheet
        cells = workbook.Worksheets[2].Cells;

        //Get the C10 cell in the third sheet
        Aspose.Cells.Cell C = cells["C10"];

        //Enter a value
        C.PutValue("page3");

        //Create a main PDF Bookmark entry object
        PdfBookmarkEntry pbeRoot = new PdfBookmarkEntry();

        //Specify its text
        pbeRoot.Text = "Sections";

        //Set the destination cell/location
        pbeRoot.Destination = p;

        //Set its sub entry array list
        pbeRoot.SubEntry = new ArrayList();

        //Create a sub PDF Bookmark entry object
        PdfBookmarkEntry subPbe1 = new PdfBookmarkEntry();

        //Specify its text
        subPbe1.Text = "Section 1";

        //Set its destination cell
        subPbe1.Destination = A;

        //Define/Create a sub Bookmark entry object of "Section A"
        PdfBookmarkEntry ssubPbe = new PdfBookmarkEntry();

        //Specify its text
        ssubPbe.Text = "Section 1.1";

        //Set its destination
        ssubPbe.Destination = D;

        //Create/Set its sub entry array list object
        subPbe1.SubEntry = new ArrayList();

        //Add the object to "Section 1"
        subPbe1.SubEntry.Add(ssubPbe);

        //Add the object to the main PDF root object
        pbeRoot.SubEntry.Add(subPbe1);

        //Create a sub PDF Bookmark entry object
        PdfBookmarkEntry subPbe2 = new PdfBookmarkEntry();

        //Specify its text
        subPbe2.Text = "Section 2";

        //Set its destination
        subPbe2.Destination = B;

        //Add the object to the main PDF root object
        pbeRoot.SubEntry.Add(subPbe2);

        //Create a sub PDF Bookmark entry object
        PdfBookmarkEntry subPbe3 = new PdfBookmarkEntry();

        //Specify its text
        subPbe3.Text = "Section 3";

        //Set its destination
        subPbe3.Destination = C;

        //Add the object to the main PDF root object
        pbeRoot.SubEntry.Add(subPbe3);

        //Set the PDF Bookmark root object
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions(SaveFormat.Pdf);
        pdfSaveOptions.Bookmark = pbeRoot;

        //Save the workbook as a PDF File
        workbook.Save(HttpContext.Current.Response, "Xls2Pdf.pdf", ContentDisposition.Attachment, pdfSaveOptions);

        //End response to avoid unneeded html after xls
        HttpContext.Current.Response.End();

    }
}
