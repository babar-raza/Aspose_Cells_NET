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
Imports Aspose.Cells.Drawing
Imports Aspose.Cells.Charts


Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for BarChart.
	''' </summary>
	Public Class BarChart
		Inherits System.Web.UI.Page
		Protected ChartTypeList As System.Web.UI.WebControls.DropDownList
		Protected WithEvents btnProcess As System.Web.UI.WebControls.Button
		Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

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
'			Me.btnProcess.Click += New System.EventHandler(Me.btnProcess_Click);
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region

		Private Sub btnProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProcess.Click
			'Initialize Workbook
			Dim workbook As New Workbook()

			'Set default font for workbook
			Dim style As Style = workbook.DefaultStyle
			style.Font.Name = "Tahoma"
			workbook.DefaultStyle = style

			'Insert Dummy Data
			CreateStaticData(workbook)

			'Apply Style on various cells
			CreateCellsFormatting(workbook)

			'Create Chart and Set Chart properties
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
			workbook.Save(HttpContext.Current.Response, "3DBarChart." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Initialize worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Hide gridlines of worksheet
			sheet.IsGridlinesVisible = False

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put value into cells
			cells.SetColumnWidth(0,13.00)
			cells("A1").PutValue("Region")
			cells("B1").PutValue("Attendance")
			cells("A2").PutValue("Providence")
			cells("B2").PutValue(120)
			cells("A3").PutValue("Philadelphia")
			cells("B3").PutValue(150)
			cells("A4").PutValue("Atlanta")
			cells("B4").PutValue(180)
			cells("A5").PutValue("Charleston")
			cells("B5").PutValue(330)
			cells("A6").PutValue("Detroit")
			cells("B6").PutValue(380)
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			Dim style1 As Style = workbook.Styles(workbook.Styles.Add())
			'Set borders
			style1.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
			style1.Font.IsBold = True
			style1.HorizontalAlignment = TextAlignmentType.Center
			style1.VerticalAlignment = TextAlignmentType.Center

			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set the width of the specified column 
			cells.SetColumnWidth(1, 13.00)

			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)

			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())
			'Copy data from another style object
			style2.Copy(style1)
			style2.Font.IsBold = False
			style2.HorizontalAlignment = TextAlignmentType.Right
			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)
			style2.Pattern = BackgroundType.Solid

			For i As Integer = 1 To 11
				If i Mod 2 <> 0 Then
					cells(i, 0).SetStyle(style2)
					cells(i, 1).SetStyle(style2)
				End If
			Next i

			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())
			style3.Copy(style2)
			'Sets foreground color
			style3.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)
			style3.Pattern = BackgroundType.Solid

			For i As Integer = 1 To 11
				If i Mod 2 = 0 Then
					cells(i, 0).SetStyle(style2)
					cells(i, 1).SetStyle(style3)
				End If
			Next i
		End Sub
		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'get worksheet index after adding new worksheet
			Dim sheetIndex As Integer = workbook.Worksheets.Add(SheetType.Chart)

			'intialize worksheet on given index
			Dim sheet As Worksheet = workbook.Worksheets(sheetIndex)

			'Set the name of worksheet
			sheet.Name = "3DBar Chart"

			'Create chart depending on selected value from ChartTypeList
			Dim chartIndex As Integer = 0
			Select Case ChartTypeList.SelectedItem.Text
				Case "CylindericalBar"
					chartIndex = sheet.Charts.Add(ChartType.CylindricalBar, 0, 0, 0, 0)
				Case "ConicalBar"
					chartIndex = sheet.Charts.Add(ChartType.ConicalBar,0,0,0,0)
				Case "PyramidBar"
					chartIndex = sheet.Charts.Add(ChartType.PyramidBar,0,0,0,0)
			End Select

			'Initialize Chart
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set properties of chart not to show Border
			chart.PlotArea.Border.IsVisible = False

			'Set properties of chart not to show legend
			chart.ShowLegend = False

			'Set properties of chart title
			chart.Title.Text = "Attendance By Region"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set properties of nseries
			chart.NSeries.Add("Sheet1!B2:B6",True)
			chart.NSeries.CategoryData = "Sheet1!A2:A6"

			'Set properties of valueaxis title
			chart.ValueAxis.Title.Text = "Attendance"
			chart.ValueAxis.Title.TextFont.Color = Color.Black
			chart.ValueAxis.Title.TextFont.IsBold = True
			chart.ValueAxis.Title.TextFont.Size = 10

			'Set properties of categoryaxis
			chart.CategoryAxis.IsPlotOrderReversed = True
		End Sub

	End Class
End Namespace
