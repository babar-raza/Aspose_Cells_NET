Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports System.Configuration
Imports System.Collections
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports Aspose.Cells

Partial Public Class Union_Intersection
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub
	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		'Call Method to create report
		CreateStaticReport()
	End Sub
	Public Sub CreateStaticReport()

		'Open template
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
		path = path.Substring(0, path.LastIndexOf("\"))
		path &= "\designer\Workbooks\BKRanges.xls"

		'Instantiate a new Workbook object.
		Dim workbook As New Workbook(path)

		'Get the first worksheet
		Dim sheet As Worksheet = workbook.Worksheets(0)


		'Get the named ranges.
		Dim ranges() As Range = workbook.Worksheets.GetNamedRanges()

		'Check whether the first range intersect the second range.
		Dim isintersect As Boolean = ranges(0).IsIntersect(ranges(1))

		'Create a style object.
		Dim style As Aspose.Cells.Style = workbook.Styles(workbook.Styles.Add())

		'Set the shading color with solid pattern type.
		style.ForegroundColor = System.Drawing.Color.Green
		style.Pattern = BackgroundType.Solid

		'Create a styleflag object.
		Dim flag As New StyleFlag()

		'Apply the cellshading.
		flag.CellShading = True

		'If first range intersects second range.
		If isintersect Then

			'Create a range by getting the intersection.
			Dim intersection As Range = ranges(0).Intersect(ranges(1))

			'Name the range.
			intersection.Name = "intersection"

			'Apply the style to the range.
			intersection.ApplyStyle(style, flag)


		End If

		'Create a style object.
		Dim style2 As Aspose.Cells.Style = workbook.Styles(workbook.Styles.Add())

		'Set the shading color with solid pattern type.
		style2.ForegroundColor = System.Drawing.Color.Yellow
		style2.Pattern = BackgroundType.Solid

		'Create a styleflag object.
		Dim flag2 As New StyleFlag()

		'Apply the cellshading.
		flag2.CellShading = True

		'Creates an arraylist.
		Dim al As New ArrayList()

		'Get the arraylist collection and apply the union operation on
		'the third and fourth ranges
		al = ranges(2).Union(ranges(3))

		'Define a range object.
		Dim union As Range

		For i As Integer = 0 To al.Count - 1

			'Get a range.
			union = CType(al(i), Range)
			'Apply the style to the range.
			union.ApplyStyle(style2, flag2)
		Next i

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "UnionAndIntersection.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "UnionAndIntersection.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()

	End Sub
End Class
