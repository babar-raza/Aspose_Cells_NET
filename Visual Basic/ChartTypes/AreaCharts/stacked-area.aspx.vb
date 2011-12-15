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
	''' Summary description for StackedArea.
	''' </summary>
	Public Class StackedArea
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
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region

		''' <summary>
		''' Create WorkBook, insert dummy data in worksheet
		''' Create chart based on the dummy data
		''' Save n xls or xlsx file format
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
		Protected Sub btnProcess_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnProcess.Click
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
			workbook.Save(HttpContext.Current.Response, "StackedArea." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			' Initialize worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the name of worksheet
			sheet.Name = "Data"

			'Set Gridlines invisible
			sheet.IsGridlinesVisible = False

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put values for rows in column A
			cells("A1").PutValue("Region")
			cells("A2").PutValue("France")
			cells("A3").PutValue("Germany")
			cells("A4").PutValue("England")

			'Put values in cells for row 1
			cells("B1").PutValue(2002)
			cells("C1").PutValue(2003)
			cells("D1").PutValue(2004)
			cells("E1").PutValue(2005)
			cells("F1").PutValue(2006)

			'put values in cells for row 2
			cells("B2").PutValue(5000)
			cells("C2").PutValue(15000)
			cells("D2").PutValue(35000)
			cells("E2").PutValue(30000)
			cells("F2").PutValue(20000)

			'put values in cells for row3
			cells("B3").PutValue(10000)
			cells("C3").PutValue(25000)
			cells("D3").PutValue(40000)
			cells("E3").PutValue(52000)
			cells("F3").PutValue(60000)

			'put values in cells for row 4
			cells("B4").PutValue(40000)
			cells("C4").PutValue(45000)
			cells("D4").PutValue(50000)
			cells("E4").PutValue(55000)
			cells("F4").PutValue(70000)
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'initialize Style 1
			Dim style1 As Style = workbook.Styles(workbook.Styles.Add())

			'Set border settings for style1
			style1.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'set Font property
			style1.Font.IsBold = True

			'Set aligment settings for Style1
			style1.HorizontalAlignment = TextAlignmentType.Center
			style1.VerticalAlignment = TextAlignmentType.Center

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Apply Style1 on first row
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)
			cells("C1").SetStyle(style1)
			cells("D1").SetStyle(style1)
			cells("E1").SetStyle(style1)
			cells("F1").SetStyle(style1)

			'Intialize Style2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set font property isBold to false
			style2.Font.IsBold = False

			'set Aligment properties of Style
			style2.HorizontalAlignment = TextAlignmentType.Right
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Pattern for Style2
			style2.Pattern = BackgroundType.Solid

			'Set Style to A2 and A4
			cells("A2").SetStyle(style2)
			cells("A4").SetStyle(style2)

			'intialize Style3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy properties from Style2
			style3.Copy(style2)

			'Set cell format
			style3.Custom = """$""#,##0"


			'Loop over the cells and set Style
			For i As Integer = 1 To 3
				If i Mod 2 <>0 Then
					cells(i, 1).SetStyle(style3)
					cells(i,2).SetStyle(style3)
					cells(i,3).SetStyle(style3)
					cells(i,4).SetStyle(style3)
					cells(i,5).SetStyle(style3)
				End If
			Next i


			'initialize Style4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy properties from style2
			style4.Copy(style2)

			'Sets foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set pattern for Style4
			style4.Pattern = BackgroundType.Solid

			'Apply Style4 to A3
			cells("A3").SetStyle(style4)

			'Initialize Style5
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy properties from Style4
			style5.Copy(style4)

			'Set cell format
			style5.Custom = """$""#,##0"

			'loop over the cells and Apply Style
			For i As Integer = 1 To 3
				If i Mod 2 = 0 Then
					cells(i,1).SetStyle(style5)
					cells(i,2).SetStyle(style5)
					cells(i,3).SetStyle(style5)
					cells(i,4).SetStyle(style5)
					cells(i,5).SetStyle(style5)
				End If
			Next i
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'get next index for Worksheet
			Dim sheetIndex As Integer = workbook.Worksheets.Add()

			'initialize worksheet on given index
			Dim sheet As Worksheet = workbook.Worksheets(sheetIndex)

			'Set the name of worksheet
			sheet.Name = "Chart"

			'Create chart, If Check box on Ui is Checked then create Area3DStacked Chart else AreaStacked
			Dim chartIndex As Integer = 0
			If CheckBoxShow3D.Checked Then
				chartIndex = sheet.Charts.Add(ChartType.Area3DStacked,1,1,25,10)
			Else
				chartIndex = sheet.Charts.Add(ChartType.AreaStacked,1,1,25,10)
			End If
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set legend position to top
			chart.Legend.Position = LegendPositionType.Top

			'if Check box on UI is checked then hide Gridlines else show them
			chart.CategoryAxis.MajorGridLines.IsVisible = False
			If CheckBoxShow3D.Checked Then
			 chart.PlotArea.Border.IsVisible = False
			End If

			'Set properties of title
			chart.Title.Text = "Total Sales "
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set properties of nseries
			chart.NSeries.Add("Data!B2:F4", False)

			'set category data from B1 to F1 Area
			chart.NSeries.CategoryData = "Data!B1:F1"

			'set visual properties for NSeries
			chart.NSeries.IsColorVaried = True

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'loop over the Chart's Nseries and Assign Name from Cell Values
			For i As Integer = 0 To chart.NSeries.Count - 1
				chart.NSeries(i).Name = cells(i+1,0).Value.ToString()
				chart.NSeries(i).Points(i).Area.ForegroundColor = Color.Red
			Next i

			'Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Year(2002-2006)"
			chart.CategoryAxis.Title.TextFont.Color = Color.Black
			chart.CategoryAxis.Title.TextFont.IsBold = True
			chart.CategoryAxis.Title.TextFont.Size = 10
			chart.CategoryAxis.AxisBetweenCategories = False
		End Sub



	End Class
End Namespace
