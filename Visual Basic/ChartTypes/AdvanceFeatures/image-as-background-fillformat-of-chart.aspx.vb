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
Imports System.IO
Imports Aspose.Cells
Imports Aspose.Cells.Drawing
Imports Aspose.Cells.Charts


Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for ImageFillFormat.
	''' </summary>
	Partial Public Class ImageFillFormat
		Inherits System.Web.UI.Page
		Protected CheckBoxShow3D As System.Web.UI.WebControls.CheckBox

		Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
			' Put user code to initialize the page here			
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
			Me.ID = "Area"

		End Sub
		#End Region

		Protected Sub btnProcess_Click(ByVal sender As Object, ByVal e As EventArgs)
			Dim workbook As New Workbook()

			'Set default font
			Dim style As Style = workbook.DefaultStyle
			style.Font.Name = "Tahoma"
			workbook.DefaultStyle = style

			CreateStaticReport(workbook)

			'Create an object of SaveFormat
			Dim saveFormat As New SaveFormat()

			'Check file format is xls
			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'Set save format optoin to xls
				saveFormat = SaveFormat.Excel97To2003
			'Check file format is xlsx
			ElseIf ddlFileVersion.SelectedItem.Value = "XLSX" Then
				'Set save format optoin to xlsx
				saveFormat = SaveFormat.Xlsx
			End If

			'Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "ImageFillFormat." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))

			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'Get First Worksheet of the Workbook
			Dim ws As Worksheet = workbook.Worksheets(0)

			'Adding data to cells
			Dim cells As Cells = ws.Cells

			'Insert cell String contents in Column A
			cells("A1").PutValue("Aspose.Cells")

			cells("A2").PutValue("Aspose.Words")

			cells("A3").PutValue("Aspose.PDF")

			'Insert Cell number contents in Column B
			cells("B1").PutValue(35)

			cells("B2").PutValue(50)

			cells("B3").PutValue(15)


			'Create a Pie Type Chart in worksheet charts collection
			Dim index As Integer = ws.Charts.Add(ChartType.Pie, 4, 1, 30, 10)

			Dim chart As Chart = ws.Charts(index)

			'Assign range of cells as charts N-Series
			chart.NSeries.Add("B1:B3", True)

			'define N-series Category Data
			chart.NSeries.CategoryData = "A1:A3"

			'Set Datalabels
			chart.NSeries(0).DataLabels.ShowPercentage = True


			'Create a Stream object AND intialize it with path to Image
			Dim fstream As New FileStream(System.Web.HttpContext.Current.Server.MapPath("~/Image/school.jpg"), FileMode.Open)

			'Read Byte Data into any Array
			Dim ImageData(fstream.Length - 1) As Byte

			'Obtain the file into the array of bytes from streams.
			fstream.Read(ImageData, 0, ImageData.Length)

			'Fillformat as Image
			chart.ChartArea.Area.FillFormat.ImageData = ImageData

			ws.AutoFitColumns()
			fstream.Close()
		End Sub
	End Class
End Namespace
