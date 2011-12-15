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
	''' Summary description for Line.
	''' </summary>
	Public Class Line
		Inherits System.Web.UI.Page
		Protected ChartTypeList As System.Web.UI.WebControls.DropDownList
		Protected NSeriesMarkerStyle As System.Web.UI.WebControls.DropDownList
		Protected NMarkBackColor As System.Web.UI.WebControls.DropDownList
		Protected NMarkForeColor As System.Web.UI.WebControls.DropDownList
		Protected NSeriesMarkSize As System.Web.UI.WebControls.DropDownList
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
			workbook.Save(HttpContext.Current.Response, "Line." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Initialize Cells object
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put string into a cells of Column A
			cells("A1").PutValue("Region")
			cells("A2").PutValue("France")
			cells("A3").PutValue("Germany")
			cells("A4").PutValue("England")

			'Put a value into a Row 1
			cells("B1").PutValue(2002)
			cells("C1").PutValue(2003)
			cells("D1").PutValue(2004)
			cells("E1").PutValue(2005)
			cells("F1").PutValue(2006)

			'Put a value into a Row 2
			cells("B2").PutValue(40000)
			cells("C2").PutValue(45000)
			cells("D2").PutValue(50000)
			cells("E2").PutValue(55000)
			cells("F2").PutValue(70000)

			'Put a value into a Row 3
			cells("B3").PutValue(10000)
			cells("C3").PutValue(25000)
			cells("D3").PutValue(40000)
			cells("E3").PutValue(52000)
			cells("F3").PutValue(60000)

			'Put a value into a Row 4
			cells("B4").PutValue(5000)
			cells("C4").PutValue(15000)
			cells("D4").PutValue(35000)
			cells("E4").PutValue(30000)
			cells("F4").PutValue(20000)
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'Initialize Style Object
			Dim style1 As Style = workbook.Styles(workbook.Styles.Add())

			'Set borders setting for Style
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

			'Set Alignments for Style
			style1.HorizontalAlignment = TextAlignmentType.Center
			style1.VerticalAlignment = TextAlignmentType.Center

			'Initalize Cells Object
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set style for Row 1
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)
			cells("C1").SetStyle(style1)
			cells("D1").SetStyle(style1)
			cells("E1").SetStyle(style1)
			cells("F1").SetStyle(style1)

			' Initialize Style 2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set IsBold Off
			style2.Font.IsBold = False

			'Set Alignment Settings for Style
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Pattern for Style
			style2.Pattern = BackgroundType.Solid

			'Apply Style2 to cells A2 and A4
			cells("A2").SetStyle(style2)
			cells("A4").SetStyle(style2)


			'Initialize Style3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'copy properties from Style2
			style3.Copy(style2)

			'Set cell format
			style3.Custom = """$""#,##0"

			'Loop cells and Set Style3
			For i As Integer = 1 To 3
				If i Mod 2 <> 0 Then
					cells(i, 1).SetStyle(style3)
					cells(i, 2).SetStyle(style3)
					cells(i, 3).SetStyle(style3)
					cells(i, 4).SetStyle(style3)
					cells(i, 5).SetStyle(style3)
				End If
			Next i

			'Initialize Style4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy properties from style2
			style4.Copy(style2)

			'Sets foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set Style Pattern
			style4.Pattern = BackgroundType.Solid

			'Apply Style4 to Cell A3
			cells("A3").SetStyle(style4)


			'Iniatalize Style5
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy Style4 properties
			style5.Copy(style4)

			'Set cell format
			style5.Custom = """$""#,##0"

			'Loop Cells ans set STyle
			For i As Integer = 1 To 3
				If i Mod 2 = 0 Then
					cells(i, 1).SetStyle(style5)
					cells(i, 2).SetStyle(style5)
					cells(i, 3).SetStyle(style5)
					cells(i, 4).SetStyle(style5)
					cells(i, 5).SetStyle(style5)
				End If
			Next i
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'Initialize Worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the name of worksheet
			sheet.Name = "Line"

			'Set Gridlines invisible
			sheet.IsGridlinesVisible = False

			'Create chart depending on the Chart Type selected from List On UI
			Dim chartIndex As Integer = 0
			Select Case ChartTypeList.SelectedItem.Text
				Case "Line"
					chartIndex = sheet.Charts.Add(ChartType.Line,5,1,29,10)
				Case "LineWithDataMarkers"
					chartIndex = sheet.Charts.Add(ChartType.LineWithDataMarkers,5,1,29,10)
			End Select

			'Initialize Chart Object
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set Chart's MajorGridLines invisib;e
			chart.CategoryAxis.MajorGridLines.IsVisible = False

			'Set Properties of chart title 
			chart.Title.Text = "Sales By Region For Years"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set Properties of nseries
			chart.NSeries.Add("B2:F4", False)

			'Set Datasource for Nseries Category
			chart.NSeries.CategoryData = "B1:F1"

			'Set Nseries color varience to True
			chart.NSeries.IsColorVaried = True


			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Loop over the cells
			For i As Integer = 0 To chart.NSeries.Count - 1
				'Set Nseriese Name from Value in cells
				chart.NSeries(i).Name = cells(i+1,0).Value.ToString()

				'Selection from dropdownlist by name NSeriesMarkerStyle
				'Set MarkerStyle depending on selection
				Select Case NSeriesMarkerStyle.SelectedItem.Text
					Case "Automatic"
						chart.NSeries(i).MarkerStyle = ChartMarkerType.Automatic
					Case "Circle"
						chart.NSeries(i).MarkerStyle = ChartMarkerType.Circle
					Case "Dash"
						chart.NSeries(i).MarkerStyle = ChartMarkerType.Dash
					Case "Diamond"
						chart.NSeries(i).MarkerStyle = ChartMarkerType.Diamond
					Case "Dot"
						chart.NSeries(i).MarkerStyle = ChartMarkerType.Dot
					Case "None"
						chart.NSeries(i).MarkerStyle = ChartMarkerType.None
					Case "Square"
						chart.NSeries(i).MarkerStyle = ChartMarkerType.Square
					Case "SquarePlus"
						chart.NSeries(i).MarkerStyle = ChartMarkerType.SquarePlus
					Case "SquareStar"
						chart.NSeries(i).MarkerStyle = ChartMarkerType.SquareStar
					Case "SquareX"
						chart.NSeries(i).MarkerStyle = ChartMarkerType.SquareX
					Case "Triangle"
						chart.NSeries(i).MarkerStyle = ChartMarkerType.Triangle
				End Select


				'Set Nseriese Marker Background and ForeGround colors
				chart.NSeries(i).MarkerBackgroundColor = Color.FromName(NMarkBackColor.SelectedItem.Text)
				chart.NSeries(i).MarkerForegroundColor = Color.FromName(NMarkForeColor.SelectedItem.Text)

				'Set NSeries Marker size 
				chart.NSeries(i).MarkerSize = Integer.Parse(NSeriesMarkSize.SelectedItem.Text)

				'Set Properties of categoryaxis title 
				chart.CategoryAxis.Title.Text = "Year(2002-2006)"
				chart.CategoryAxis.Title.TextFont.Color = Color.Black
				chart.CategoryAxis.Title.TextFont.IsBold = True
				chart.CategoryAxis.Title.TextFont.Size = 10

				'Set the legend position to Top
				chart.Legend.Position = LegendPositionType.Top
			Next i
		End Sub

	End Class
End Namespace
