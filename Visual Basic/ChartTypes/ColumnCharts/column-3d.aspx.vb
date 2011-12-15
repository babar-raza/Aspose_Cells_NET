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
	''' Summary description for Column3D.
	''' </summary>
	Public Class Column3D
		Inherits System.Web.UI.Page
		Protected ColumnType As System.Web.UI.WebControls.DropDownList
		Protected WallsColor As System.Web.UI.WebControls.DropDownList
		Protected FloorColor As System.Web.UI.WebControls.DropDownList
		Protected Rotation As System.Web.UI.WebControls.DropDownList
		Protected Elevation As System.Web.UI.WebControls.DropDownList
		Protected DepthPercent As System.Web.UI.WebControls.DropDownList
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
			workbook.Save(HttpContext.Current.Response, "3DColumn." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put string values in row cells of column 1
			cells("A1").PutValue("Region")
			cells("A2").PutValue("France")
			cells("A3").PutValue("Germany")
			cells("A4").PutValue("England")

			'Put values in row cells of column 3
			cells("B1").PutValue("Marketing Costs")
			cells("B2").PutValue(70000)
			cells("B3").PutValue(55000)
			cells("B4").PutValue(30000)
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'Initialize Style1
			Dim style1 As Style = workbook.Styles(workbook.Styles.Add())

			'Set border for Style1
			style1.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'set Font Property IsBold to true
			style1.Font.IsBold = True

			'Set Alignment of Style
			style1.HorizontalAlignment = TextAlignmentType.Center
			style1.VerticalAlignment = TextAlignmentType.Center

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set the width of the specified column 
			cells.SetColumnWidth(1,15)

			'Apply Style to A1 and B1
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)

			'Initialize Style2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set Font IsBold Property to False
			style2.Font.IsBold = False

			'Set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Style Pattern
			style2.Pattern = BackgroundType.Solid

			'Apply Style to A2 and A4
			cells("A2").SetStyle(style2)
			cells("A4").SetStyle(style2)

			'Initialize Style3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy the properties from Style2
			style3.Copy(style2)

			'Set cell format
			style3.Custom = """$""#,##0"

			'Apply Style to Cell B2 and B4
			cells("B2").SetStyle(style3)
			cells("B4").SetStyle(style3)

			'initialize Style4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy the properties of Style2
			style4.Copy(style2)

			'Sets foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set Styte Pattern
			style4.Pattern = BackgroundType.Solid

			'Apply Style on cell A3
			cells("A3").SetStyle(style4)


			'initialize Style5
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy the properties of Style4
			style5.Copy(style4)

			'Set cell format
			style5.Custom = """$""#,##0"

			'Set Style on cell B3
			cells("B3").SetStyle(style5)
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'Initialize Worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the name of the worksheet. 
			sheet.Name = "3D Column"

			'Set Gridlines invisible
			sheet.IsGridlinesVisible = False

			'Create chart
			Dim chartIndex As Integer = 0

			'Select Chart Type based on Values in Column Type drop down List
			Select Case ColumnType.SelectedItem.Text
				Case "Column3D"
					chartIndex = sheet.Charts.Add(ChartType.Column3D,5,1,29,10)
				Case "Column3DClustered"
					chartIndex = sheet.Charts.Add(ChartType.Column3DClustered,5,1,29,10)
				Case "Column3DStacked"
					chartIndex = sheet.Charts.Add(ChartType.Column3DStacked,5,1,29,10)
			End Select

			'Initialize Chart
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set properties of chart 
			chart.CategoryAxis.MajorGridLines.IsVisible = False

			'Set properties of nseries
			chart.NSeries.Add("B2:B4", True)
			chart.NSeries.CategoryData = "A2:A4"
			chart.NSeries.IsColorVaried = True

			'Set properties of chart title 
			chart.Title.Text = "Marketing Costs by Region"
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.Size = 12

			'Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Region"
			chart.CategoryAxis.Title.TextFont.Color = Color.Black
			chart.CategoryAxis.Title.TextFont.IsBold = True
			chart.CategoryAxis.Title.TextHorizontalAlignment = TextAlignmentType.Center
			chart.CategoryAxis.Title.TextFont.Size = 10

			'Set properties of valueaxis title 
			chart.ValueAxis.Title.Text = "In Thousands"
			chart.ValueAxis.Title.TextFont.Color = Color.Black
			chart.ValueAxis.Title.TextFont.IsBold = True
			chart.ValueAxis.Title.TextFont.Size = 10
			chart.ValueAxis.Title.RotationAngle = 90

			'Set the legend position  to Top
			chart.Legend.Position = LegendPositionType.Top

			'Set Borders of chart invisible
			chart.PlotArea.Border.IsVisible = False

			'Set properties of chart based on values in controls from UI
			chart.Walls.ForegroundColor = Color.FromName(WallsColor.SelectedItem.Text)
			chart.Floor.ForegroundColor = Color.FromName(FloorColor.SelectedItem.Text)
			chart.RotationAngle = Integer.Parse(Rotation.SelectedItem.Text)
			chart.Elevation = Integer.Parse(Elevation.SelectedItem.Text)
			chart.DepthPercent = Integer.Parse(DepthPercent.SelectedItem.Text)
		End Sub

	End Class
End Namespace
