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
	''' Summary description for StackedColumn.
	''' </summary>
	Public Class StackedColumn
		Inherits System.Web.UI.Page
		Protected checkBoxShow3D As System.Web.UI.WebControls.CheckBox
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
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region

		Protected Sub btnProcess_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnProcess.Click
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
			workbook.Save(HttpContext.Current.Response, "PercentStackedColumn." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Initialize Worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the name of worksheet
			sheet.Name = "Data"

			'Set worksheets Gridlines to invisible
			sheet.IsGridlinesVisible = False


			'initialize cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put values in rows for Column1
			cells("A1").PutValue("Year")
			cells("A2").PutValue(2004)
			cells("A3").PutValue(2005)
			cells("A4").PutValue(2006)

			'Put values in rows for Column2
			cells("B2").PutValue(20000)
			cells("B3").PutValue(40000)
			cells("B4").PutValue(40000)

			'Put values in rows for Column3
			cells("C2").PutValue(30000)
			cells("C3").PutValue(20000)
			cells("C4").PutValue(50000)

			'Put value in CElls B1 and B2 for ROw Headers
			cells("B1").PutValue("Product1")
			cells("C1").PutValue("Product2")
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'initialize Style1
			Dim style1 As Style = workbook.Styles(workbook.Styles.Add())
			'Set Boder settings for style1
			style1.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'Set Font Style
			style1.Font.IsBold = True

			'Set Alignments for Style2
			style1.HorizontalAlignment = TextAlignmentType.Center
			style1.VerticalAlignment = TextAlignmentType.Center

			'//initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Apply Style to Row Headers
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)
			cells("C1").SetStyle(style1)

			'initialize Style2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)
			style2.Font.IsBold = False
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Style Pattern
			style2.Pattern = BackgroundType.Solid

			'Set Style to A2 and A4
			cells("A2").SetStyle(style2)
			cells("A4").SetStyle(style2)

			'initialize Style3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style3.Copy(style2)

			'Set cell format
			style3.Custom = """$""#,##0"

			'Set Style to Cells B, C2, B4 and C4
			cells("B2").SetStyle(style3)
			cells("C2").SetStyle(style3)
			cells("B4").SetStyle(style3)
			cells("C4").SetStyle(style3)

			'initialize Style4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style4.Copy(style2)

			'Sets foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set Style Pattern
			style4.Pattern = BackgroundType.Solid

			'Set Style to cell A3
			cells("A3").SetStyle(style4)

			'initialize Style5
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style5.Copy(style4)

			'Set cell format
			style5.Custom = """$""#,##0"

			'Set Style to Cells B3 and C3
			cells("B3").SetStyle(style5)
			cells("C3").SetStyle(style5)
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'get next index for new worksheet
			Dim sheetIndex As Integer = workbook.Worksheets.Add()

			'Initalize worksheet for given index
			Dim sheet As Worksheet = workbook.Worksheets(sheetIndex)

			'Set the name of worksheet
			sheet.Name = "Chart"

			'Create Chart
			Dim chartIndex As Integer = 0
			' Show as 2d or 3d depending on the state of check Box on UI
			If checkBoxShow3D.Checked Then
				chartIndex = sheet.Charts.Add(ChartType.Column3DStacked, 1, 1, 25, 10)
			Else
				chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 1, 1, 25, 10)
			End If
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set properies to chart
			chart.CategoryAxis.MajorGridLines.IsVisible =False
			If checkBoxShow3D.Checked Then
				chart.PlotArea.Border.IsVisible = False
			End If

			'Set properies to chart title
			chart.Title.Text = "Product  Sales"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set properies to nseries
			chart.NSeries.CategoryData = "Data!A2:A4"

			'Set NSeries Data
			chart.NSeries.Add("Data!B2:C4", True)

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells


			'Iterate over the NSeries and assign it name from values cells
			For i As Integer = 0 To chart.NSeries.Count - 1
				chart.NSeries(i).Name = cells(0,i + 1).Value.ToString()
			Next i

			'Set properies to categoryaxis
			chart.CategoryAxis.Title.Text = "Year(2004-2006)"
			chart.CategoryAxis.Title.TextFont.Color = Color.Black
			chart.CategoryAxis.Title.TextFont.Size = 10
			chart.CategoryAxis.Title.TextFont.IsBold = True

			'Set the legend position To Tip
			chart.Legend.Position = LegendPositionType.Top
		End Sub
	End Class
End Namespace
