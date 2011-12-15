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
using System.Data.OleDb;

namespace Aspose.Cells.Demos.SmartMarker
{
    public partial class designer : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnProcess_Click(object sender, EventArgs e)
        {
            //Open the template file through streams
            string path = MapPath(".");
            path = path.Substring(0, path.LastIndexOf("\\")) + "\\Designer\\SmartMarkerDesigner.xls";
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            byte[] data = new byte[fs.Length];
            fs.Read(data, 0, data.Length);
            fs.Close();

            //Open/Save the template file through Response object
            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("content-disposition", "attachment;  filename=SmartMarkerDesigner.xls");
            Response.BinaryWrite(data);
            Response.End();
        }

    }
}


