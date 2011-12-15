Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Web
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports Aspose.Cells

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for Merge/UnMerge Cells
	''' </summary>
	Public Class MergeUnMergeCells
		Inherits System.Web.UI.Page
		Protected WithEvents btnMerge As System.Web.UI.WebControls.Button
		Protected WithEvents btnUnMerge As System.Web.UI.WebControls.Button

		Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			' Put user code to initialize the page here
		End Sub

		#Region "Web Form Designer generated code"
		Overrides Protected Sub OnInit(ByVal e As EventArgs)
			'
			' CODEGEN: This call is required by the ASP.NET Web Form Designer.
			'
			 If Context IsNot Nothing AndAlso Context.Session IsNot Nothing Then
				InitializeComponent()
				MyBase.OnInit(e)
			 End If
		End Sub

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
'			Me.btnMerge.Click += New System.EventHandler(Me.btnMerge_Click);
'			Me.btnUnMerge.Click += New System.EventHandler(Me.btnUnMerge_Click);
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region

		Private Sub btnMerge_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMerge.Click
			'Create a Workbook.
			Dim wbk As New Aspose.Cells.Workbook()

			'Create a Worksheet and get the first sheet.
			Dim worksheet As Aspose.Cells.Worksheet = wbk.Worksheets(0)

			'Create a Cells object ot fetch all the cells.
			Dim cells As Aspose.Cells.Cells = worksheet.Cells

			'Merge some Cells (C6:E7) into a single C6 Cell.
			cells.Merge(5, 2, 2, 3)

			'Input data into C6 Cell.
			worksheet.Cells(5, 2).PutValue("This is my value")

			'Create a Style object to fetch the Style of C6 Cell.
			Dim style As Aspose.Cells.Style = worksheet.Cells(5, 2).GetStyle()

			'Create a Font object
			Dim font As Aspose.Cells.Font = style.Font

			'Set the name.
			font.Name = "Times New Roman"

			'Set the font size.
			font.Size = 18

			'Set the font color
			font.Color = Color.Blue

			'Bold the text
			font.IsBold = True

			'Make it italic
			font.IsItalic = True

			'Set the backgrond color of C6 Cell to Red
			style.ForegroundColor = Color.Red

			style.Pattern = BackgroundType.Solid

			'Apply the Style to C6 Cell.
			cells(5, 2).SetStyle(style)

			'Save the excel file
			wbk.Save(HttpContext.Current.Response,"MergeCells.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))

			' End response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub btnUnMerge_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUnMerge.Click
			'Create a Workbook.
			Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
			path = path.Substring(0, path.LastIndexOf("\"))
			path &= "\designer\Workbooks\MergeCells.xls"

			Dim wbk As New Aspose.Cells.Workbook(path)


			'Create a Worksheet and get the first sheet.
			Dim worksheet As Aspose.Cells.Worksheet = wbk.Worksheets(0)

			'Create a Cells object ot fetch all the cells.
			Dim cells As Aspose.Cells.Cells = worksheet.Cells

			'Unmerge the cells.
			cells.UnMerge(5, 2, 2, 3)

			'Save the excel file
			wbk.Save(HttpContext.Current.Response, "UnMergeCells.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))

			' End response to avoid unneeded html after xls
			Response.End()
		End Sub
	End Class
End Namespace


