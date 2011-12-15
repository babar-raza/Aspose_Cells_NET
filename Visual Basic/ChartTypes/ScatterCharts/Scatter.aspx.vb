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
	''' Summary description for Scatter.
	''' </summary>
	Public Class Scatter
		Inherits System.Web.UI.Page
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
			Me.ID = "Scatter"
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
			workbook.Save(HttpContext.Current.Response, "Scatter." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Initialize Worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'name Worksheet
			sheet.Name = "Data"

			'Set Worksheet's gridlines invisible
			sheet.IsGridlinesVisible = False

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put string in cells to make Column Header
			cells("A1").PutValue("Daily Rainfall")
			cells("B1").PutValue("Particulate")

			'Put values to make rows
			cells("A2").PutValue(1.9)
			cells("B2").PutValue(137)
			cells("A3").PutValue(3.6)
			cells("B3").PutValue(128)
			cells("A4").PutValue(4.1)
			cells("B4").PutValue(122)
			cells("A5").PutValue(4.3)
			cells("B5").PutValue(117)
			cells("A6").PutValue(5)
			cells("B6").PutValue(114)
			cells("A7").PutValue(5.4)
			cells("B7").PutValue(114)
			cells("A8").PutValue(5.7)
			cells("B8").PutValue(112)
			cells("A9").PutValue(5.9)
			cells("B9").PutValue(110)
			cells("A10").PutValue(7.3)
			cells("B10").PutValue(104)
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'Initialize Style1
			Dim style1 As Style = workbook.Styles(workbook.Styles.Add())

			'Set border style
			style1.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'Set Font to Bold
			style1.Font.IsBold = True

			'Set Style Alignment
			style1.HorizontalAlignment = TextAlignmentType.Center

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set the width of the specified column 
			cells.SetColumnWidth(0, 12)
			cells.SetColumnWidth(1, 10)

			'Set Style for Column Header 
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)

			'Initialize Style2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set Font to Not Bold
			style2.Font.IsBold = False

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Style Patern
			style2.Pattern = BackgroundType.Solid

			'Set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right

			'loop over the cells and set Style
			For i As Integer = 1 To 9
				If i Mod 2 <> 0 Then
					cells(i, 0).SetStyle(style2)
					cells(i, 1).SetStyle(style2)
				End If
			Next i

			'Initialize Style3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style3.Copy(style2)

			'Set foreground color
			style3.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set Style pattern
			style3.Pattern = BackgroundType.Solid

			'loop over the cells and set Style
			For i As Integer = 1 To 9
				If i Mod 2 = 0 Then
					cells(i, 0).SetStyle(style2)
					cells(i, 1).SetStyle(style3)
				End If
			Next i
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'Initialize Worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the name of worksheet
			sheet.Name = "Scatter"

			'Set Worksheet's GridLines invisible
			sheet.IsGridlinesVisible = False

			'Create chart of Type Scatter
			Dim chartIndex As Integer = sheet.Charts.Add(ChartType.Scatter,1,3,25,12)

			'Initialize Chart
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set properties of chart
			chart.CategoryAxis.MajorGridLines.IsVisible = False

			'Set Legend Position of chart
			chart.Legend.Position = LegendPositionType.Top

			'Set properties of chart title
			chart.Title.Text = "Scatter Chart:Particulate Levels in Rainfall"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set properties of nseries
			chart.NSeries.Add ("B2:B10",True)
			chart.NSeries(0).XValues = "A2:A10"

			'Loop over the NSeries and set Name
			For i As Integer = 0 To chart.NSeries.Count - 1
				chart.NSeries(i).Name = "Particulate"
			Next i

			Dim cells As Cells = workbook.Worksheets(0).Cells
			'Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = cells("A1").Value.ToString()
			chart.CategoryAxis.Title.TextFont.Color = Color.Black
			chart.CategoryAxis.Title.TextFont.IsBold = True
			chart.CategoryAxis.Title.TextFont.Size = 10

			'Set properties of valueaxis title
			chart.ValueAxis.Title.Text = cells("B1").Value.ToString()
			chart.ValueAxis.Title.TextFont.Color = Color.Black
			chart.ValueAxis.Title.TextFont.IsBold = True
			chart.ValueAxis.Title.TextFont.Size = 10
			chart.ValueAxis.Title.RotationAngle = 90
		End Sub

	End Class
End Namespace
