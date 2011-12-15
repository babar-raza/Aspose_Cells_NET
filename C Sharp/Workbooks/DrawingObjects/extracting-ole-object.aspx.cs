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
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

public partial class Workbooks_DrawingObjects_ExtractingOleObject : System.Web.UI.Page
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
        //Open template from path
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\OleFile.xls";

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(path);

        //Get the OleObject Collection in the first worksheet.
        OleObjectCollection oles = workbook.Worksheets[0].OleObjects;

        //Loop through all the oleobjects and extract each object in the worksheet.
        for (int i = 0; i < oles.Count; i++)
        {
            //Create Ole Object and Initialize it with i Item in collection
            OleObject ole = oles[i];

            //Specify the output filename.
            string fileName = "outOle" + i + ".";

            //Specify each file format based on the oleobject format type.
            switch (ole.FileType)
            {

                case OleFileType.Doc:
                    fileName += "doc";
                    break;

                case OleFileType.Xls:
                    fileName += "Xls";
                    break;

                case OleFileType.Ppt:
                    fileName += "Ppt";
                    break;

                case OleFileType.Pdf:
                    fileName += "Pdf";
                    break;

                case OleFileType.Unknown:
                    fileName += "Jpg";
                    break;

                default:
                    //........
                    break;
            }


            //Save the oleobject as a new excel file if the object type is xls.
            if (ole.FileType == OleFileType.Xls)
            {
                //Create MemoryStream
                MemoryStream ms = new MemoryStream();

                //Write OleObject to Memory Stream 
                ms.Write(ole.ObjectData, 0, ole.ObjectData.Length);

                //Ctreate WorkBook from MemoryStream
                Workbook oleBook = new Workbook(ms);

                //Hide all worksheets from workbook
                oleBook.Worksheets.IsHidden = false;

                //Saving the Excel file
                oleBook.Save(HttpContext.Current.Response,"OleObect.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));

            }

           //Create the files based on the oleobject format types.                
            else
            {
                
                //FileStream fs = File.Create(fileName);
                HttpContext.Current.Response.Clear();
                HttpContext.Current.Response.ContentType = "image/jpg";
                HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=OleFile.jpg");
                HttpContext.Current.Response.OutputStream.Write(ole.ObjectData, 0, ole.ObjectData.Length);

            }

        }               
         // End response to avoid unneeded html after xls
        HttpContext.Current.Response.End();
    }
}
