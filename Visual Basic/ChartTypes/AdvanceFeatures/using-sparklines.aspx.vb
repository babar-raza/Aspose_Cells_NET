Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Configuration
Imports System.Collections
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports Aspose.Cells
Imports Aspose.Cells.Charts
Imports System.Drawing

Namespace Aspose.Cells.Demos
	Partial Public Class UsingSparklines
		Inherits System.Web.UI.Page
		Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
		#Region "Web Form Designer generated code"
		Overrides Protected Sub OnInit(ByVal e As EventArgs)
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
			'this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region
		Protected Sub btnProcess_Click(ByVal sender As Object, ByVal e As EventArgs)
			CreateStaticReport()
		End Sub

		Protected Sub CreateStaticReport()
			'Intiaalize workbook with xlsx file format
			Dim workbook As New Workbook()

			'Clear workbook's worksheets
			workbook.Worksheets.Clear()

			'Insert new Worksheet in workbook and name it "New"
			Dim worksheet As Worksheet = workbook.Worksheets.Add("New")

			'Insert dummy data in A8, A9 and A10 cells
			worksheet.Cells("A8").PutValue(34)
			worksheet.Cells("A9").PutValue(50)
			worksheet.Cells("A10").PutValue(34)

			'Intialize Cell Area
			Dim cellArea As New CellArea()

			'Assign Cell Area boundaries
			cellArea.StartColumn = 0
			cellArea.EndColumn = 0
			cellArea.StartRow = 0
			cellArea.EndRow = 0

			'Add new Sparklines in worksheet's sparlines collection and Assign the area for it
			Dim index As Integer = worksheet.SparklineGroupCollection.Add(SparklineType.Column, worksheet.Name & "!A8:A10", True, cellArea)

			'Initalize Sparklines Group
			Dim group As SparklineGroup = worksheet.SparklineGroupCollection(index)


			' change the color of the series if need
			Dim cellColor As CellsColor = workbook.CreateCellsColor()
			cellColor.Color = Color.Orange

			'Asign the group series color
			group.SeriesColor = cellColor

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "SparkLines.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "SparkLines.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub
	End Class
End Namespace
