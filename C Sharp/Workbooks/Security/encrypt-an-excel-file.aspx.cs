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

public partial class EncryptingFile : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void btnExecute_Click(object sender, EventArgs e)
    {
        CreateStaticReport();
    }

    public void CreateStaticReport()
    {
        //Open template.
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\book1.xls";

        //Instantiate a new Workbook object.
        Workbook workbook = new Workbook(path);

        //Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
        
        //Use this line if you want to specify XOR Encrytion type.
        //workbook.SetEncryptionOptions(EncryptionType.XOR, 40);
        
        //Password protect the file.
        workbook.Settings.Password = "007";
       
        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "EncryptedBook.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "EncryptedBook.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();   


    }
}
