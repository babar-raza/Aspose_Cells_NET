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
	''' Summary description for StackedBar.
	''' </summary>
	Public Class StackedBar
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
			workbook.Save(HttpContext.Current.Response, "StackedBar." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Set the name of worksheet
			sheet.Name = "Data"
			sheet.IsGridlinesVisible = False

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put string values in row cells of column 1
			cells("A1").PutValue("Region")
			cells("A2").PutValue("France")
			cells("A3").PutValue("Germany")
			cells("A4").PutValue("English")
			cells("A5").PutValue("Italy")

			'Put Number values in row cells of column 2
			cells("B2").PutValue(25000)
			cells("B3").PutValue(15000)
			cells("B4").PutValue(30000)
			cells("B5").PutValue(20000)

			'Put Number values in row cells of column 3
			cells("C2").PutValue(20000)
			cells("C3").PutValue(15000)
			cells("C4").PutValue(25000)
			cells("C5").PutValue(30000)

			'Put Number values in row cells of column 4
			cells("D2").PutValue(30000)
			cells("D3").PutValue(32000)
			cells("D4").PutValue(15000)
			cells("D5").PutValue(10000)

			'Put string values in cells B1, C1, D1
			cells("B1").PutValue("Apple")
			cells("C1").PutValue("Orange")
			cells("D1").PutValue("Banana")
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'Intialize Style1
			Dim style1 As Style = workbook.Styles(workbook.Styles.Add())

			'Set border settings for Style1
			style1.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'Set Font property IsBold to True
			style1.Font.IsBold = True

			'Set Style Alignment
			style1.HorizontalAlignment = TextAlignmentType.Center

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set Style for First Row
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)
			cells("C1").SetStyle(style1)
			cells("D1").SetStyle(style1)

			'Initialize Style 2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set Font IsBold Property to false
			style2.Font.IsBold = False

			'Set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Style Pattern
			style2.Pattern = BackgroundType.Solid

			'Set Style of cell A2 and A4
			cells("A2").SetStyle(style2)
			cells("A4").SetStyle(style2)

			'Initialize Style3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy properties from Style2
			style3.Copy(style2)

			'Set cell format
			style3.Custom = """$""#,##0"

			'Loop Over the cells and set Style
			For i As Integer = 1 To 4
				If i Mod 2 <>0 Then
					cells(i,1).SetStyle(style3)
					cells(i,2).SetStyle(style3)
					cells(i,3).SetStyle(style3)
				End If
			Next i

			'Initialize Style4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy the Properties of Style2
			style4.Copy(style2)

			'Set foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set Style Pattern
			style4.Pattern = BackgroundType.Solid

			'Apply Style to cells A3 and A5
			cells("A3").SetStyle(style4)
			cells("A5").SetStyle(style4)

			'Initialize Style5
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy the properties of Style4
			style5.Copy(style4)

			'Set cell format
			style5.Custom = """$""#,##0"

			'Loop over the cells as set the Style
			For i As Integer = 1 To 4
				If i Mod 2 = 0 Then
					cells(i,1).SetStyle(style5)
					cells(i,2).SetStyle(style5)
					cells(i,3).SetStyle(style5)
				End If
			Next i
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'get the next index for worksheets in workbook
			Dim sheetIndex As Integer = workbook.Worksheets.Add()

			'Initialize worksheet from given index
			Dim sheet As Worksheet = workbook.Worksheets(sheetIndex)

			'Set the name of worksheet
			sheet.Name = "Chart"

			'Create chart
			Dim indexChart As Integer = 0
			'Create chart, If Check box on Ui is Checked then create Bar3DStacked Chart else BarStacked Chart
			If CheckBoxShow3D.Checked Then
				indexChart = sheet.Charts.Add (ChartType.Bar3DStacked,1,1,21,10)
			Else
				indexChart = sheet.Charts.Add (ChartType.BarStacked,1,1,21,10)
			End If
			Dim chart As Chart = sheet.Charts(indexChart)

			'Set properties of chart and hide gridLines based upon the state od check box on UI
			If CheckBoxShow3D.Checked Then
				chart.PlotArea.Border.IsVisible = False
			End If
			chart.CategoryAxis.MajorGridLines.IsVisible = False

			'Set properties of chart title
			chart.Title.Text = "Fruit Sales By Region"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set properties of nseries
			chart.NSeries.Add ("Data!B2:D5", True)
			chart.NSeries.CategoryData = "Data!A2:A5"
			chart.NSeries.IsColorVaried = True

			'Initalize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'loop over the Chart's Nseries and Assign Name from Cell Values
			For i As Integer = 0 To chart.NSeries.Count - 1
				chart.NSeries(i).Name = cells(0,i+1).Value.ToString()
			Next i

			'Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Region"
			chart.CategoryAxis.Title.TextFont.Color = Color.Black
			chart.CategoryAxis.Title.TextFont.IsBold = True
			chart.CategoryAxis.Title.TextFont.Size = 10
			chart.CategoryAxis.Title.RotationAngle = 90

			'Set properties of legend
			chart.Legend.Position = LegendPositionType.Top
		End Sub
	End Class
End Namespace
