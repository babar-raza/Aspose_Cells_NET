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
	''' Summary description for AddChartDataTable.
	''' </summary>
	Public Class AddChartDataTable
		Inherits System.Web.UI.Page
		Protected CheckBoxShow3D As System.Web.UI.WebControls.CheckBox
		Protected WithEvents btnProcess As System.Web.UI.WebControls.Button
		Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

		Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
'			Me.btnProcess.Click += New System.EventHandler(Me.btnProcess_Click);
			Me.ID = "Area"
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region

		Protected Sub btnProcess_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnProcess.Click
			Dim workbook As New Workbook()

			'Set default font
			Dim style As Style = workbook.DefaultStyle
			style.Font.Name = "Tahoma"
			workbook.DefaultStyle = style

			'Insert dummy data in worksheet and Create ChartDataTable of that data
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
			workbook.Save(HttpContext.Current.Response, "AddChartDataTable." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))

			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Adding a sample value to "A1" cell
			worksheet.Cells("A1").PutValue(50)

			'Adding a sample value to "A2" cell
			worksheet.Cells("A2").PutValue(100)

			'Adding a sample value to "A3" cell
			worksheet.Cells("A3").PutValue(150)

			'Adding a sample value to "B1" cell
			worksheet.Cells("B1").PutValue(60)

			'Adding a sample value to "B2" cell
			worksheet.Cells("B2").PutValue(32)

			'Adding a sample value to "B3" cell
			worksheet.Cells("B3").PutValue(50)

			'Adding a chart to the worksheet
			Dim chartIndex As Integer = worksheet.Charts.Add(ChartType.Column, 5, 0, 25, 10)

			'Accessing the instance of the newly added chart
			Dim chart As Chart = worksheet.Charts(chartIndex)

			'Adding NSeries (chart data source) to the chart ranging from "A1" cell to "B3"
			chart.NSeries.Add("A1:B3", True)

			'Show the Data Table with the chart
			chart.ShowDataTable = True

			'Getting Chart Table
			Dim chartTable As ChartDataTable = chart.ChartDataTable

			'Setting Chart Table Font Color
			chartTable.Font.Color = System.Drawing.Color.Red

			'Setting Legend Key Visibility
			chartTable.ShowLegendKey = False

			'Setting Chart Table's border Color
			chartTable.Border.Color = Color.Blue

		End Sub
	End Class
End Namespace
