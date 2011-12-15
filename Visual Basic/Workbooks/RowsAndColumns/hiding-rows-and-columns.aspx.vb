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
	''' Summary description for HidingRowsAndColumns.
	''' </summary>
	Public Class HidingRowsAndColumns
		Inherits System.Web.UI.Page
		Protected WithEvents Button1 As System.Web.UI.WebControls.Button
		Protected WithEvents Button2 As System.Web.UI.WebControls.Button

		Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			' Put user code to initialize the page here
		End Sub

		#Region "Web Form Designer generated code"
		Overrides Protected Sub OnInit(ByVal e As EventArgs)
			'
			' CODEGEN: This call is required by the ASP.NET Web Form Designer.
			'
			InitializeComponent()
			MyBase.OnInit(e)
		End Sub

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
'			Me.Button1.Click += New System.EventHandler(Me.Button1_Click);
'			Me.Button2.Click += New System.EventHandler(Me.Button2_Click);
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region

		Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
			Dim workbook As New Workbook()

			Dim sheet As Worksheet = workbook.Worksheets(0)

			CreateStaticData(workbook)

			'Unhide the 3rd row and setting its height to 13.5
			sheet.Cells.UnhideRow(2, 13.5)
			'Unhide the 2nd column and setting its width to 15
			sheet.Cells.UnhideColumn(1, 15)

			workbook.Save(HttpContext.Current.Response, "DisplayRowsColumns.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))

			' End response to avoid unneeded html after xls
			Response.End()


		End Sub

		Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
			Dim workbook As New Workbook()

			Dim sheet As Worksheet = workbook.Worksheets(0)

			CreateStaticData(workbook)

			'Hide the 3rd row of the worksheet
			sheet.Cells.HideRow(2)
			'Hide the 2nd column of the worksheet
			sheet.Cells.HideColumn(1)

			workbook.Save(HttpContext.Current.Response, "HideRowsColumns.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))

			' End response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Set default font
			Dim style As Style = workbook.DefaultStyle
			style.Font.Name = "Tahoma"
			workbook.DefaultStyle = style

			Dim cells As Cells = workbook.Worksheets(0).Cells
			'Put a value into a cell
			cells("A1").PutValue("Year")
			cells("A2").PutValue(2005)
			cells("A3").PutValue(2006)

			cells("B1").PutValue("No. of Employees")
			cells("B2").PutValue(98)
			cells("B3").PutValue(113)
		End Sub


	End Class
End Namespace
