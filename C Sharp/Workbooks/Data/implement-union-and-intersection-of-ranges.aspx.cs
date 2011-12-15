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

public partial class Union_Intersection : System.Web.UI.Page
{
    protected System.Web.UI.WebControls.DropDownList ddlFileVersion;

    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void btnExecute_Click(object sender, EventArgs e)
    {
        //Call Method to create report
        CreateStaticReport();
    }
    public void CreateStaticReport()
    {

        //Open template
        string path = System.Web.HttpContext.Current.Server.MapPath("~");
        path = path.Substring(0, path.LastIndexOf("\\"));
        path += @"\designer\Workbooks\BKRanges.xls";

        //Instantiate a new Workbook object.
        Workbook workbook = new Workbook(path);

        //Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

       
        //Get the named ranges.
        Range[] ranges = workbook.Worksheets.GetNamedRanges();

        //Check whether the first range intersect the second range.
        bool isintersect = ranges[0].IsIntersect(ranges[1]);
        
        //Create a style object.
        Aspose.Cells.Style style = workbook.Styles[workbook.Styles.Add()];
        
        //Set the shading color with solid pattern type.
        style.ForegroundColor = System.Drawing.Color.Green;
        style.Pattern = BackgroundType.Solid;
        
        //Create a styleflag object.
        StyleFlag flag = new StyleFlag();
        
        //Apply the cellshading.
        flag.CellShading = true;

        //If first range intersects second range.
        if (isintersect)
        {

            //Create a range by getting the intersection.
            Range intersection = ranges[0].Intersect(ranges[1]);
            
            //Name the range.
            intersection.Name = "intersection";
            
            //Apply the style to the range.
            intersection.ApplyStyle(style, flag);

            
        }

        //Create a style object.
        Aspose.Cells.Style style2 = workbook.Styles[workbook.Styles.Add()];
        
        //Set the shading color with solid pattern type.
        style2.ForegroundColor = System.Drawing.Color.Yellow;
        style2.Pattern = BackgroundType.Solid;
        
        //Create a styleflag object.
        StyleFlag flag2 = new StyleFlag();
        
        //Apply the cellshading.
        flag2.CellShading = true;
        
        //Creates an arraylist.
        ArrayList al = new ArrayList();
        
        //Get the arraylist collection and apply the union operation on
        //the third and fourth ranges
        al = ranges[2].Union(ranges[3]);

        //Define a range object.
        Range union;
        
        for (int i = 0; i < al.Count; i++)
        {

            //Get a range.
            union = (Range)al[i];
            //Apply the style to the range.
            union.ApplyStyle(style2, flag2);
        }
        
        if (ddlFileVersion.SelectedItem.Value == "XLS")
        {
            ////Save file and send to client browser using selected format
            workbook.Save(HttpContext.Current.Response, "UnionAndIntersection.xls", ContentDisposition.Attachment, new XlsSaveOptions(SaveFormat.Excel97To2003));
        }
        else
        {
            workbook.Save(HttpContext.Current.Response, "UnionAndIntersection.xlsx", ContentDisposition.Attachment, new OoxmlSaveOptions(SaveFormat.Xlsx));
        }

        //end response to avoid unneeded html
        HttpContext.Current.Response.End();      

    }
}
