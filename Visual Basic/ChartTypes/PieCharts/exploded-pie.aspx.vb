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
	''' Summary description for ExplodedPie.
	''' </summary>
	Public Class ExplodedPie
		Inherits System.Web.UI.Page
		Protected WithEvents btnProcess As System.Web.UI.WebControls.Button
		Protected CheckShow3D As System.Web.UI.WebControls.CheckBox
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
			workbook.Save(HttpContext.Current.Response, "ExplodedPie." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Initialize Worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the name of worksheet
			sheet.Name = "Data"

			'Set GridLines invisible
			sheet.IsGridlinesVisible = False

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put values for row cells of Column 1
			cells("A1").PutValue("Region")
			cells("A2").PutValue("France")
			cells("A3").PutValue("Germany")
			cells("A4").PutValue("England")
			cells("A5").PutValue("Sweden")
			cells("A6").PutValue("Italy")
			cells("A7").PutValue("Spain")
			cells("A8").PutValue("Portugal")

			'Put values for row cells of Column 2
			cells("B1").PutValue("Sale")
			cells("B2").PutValue(70000)
			cells("B3").PutValue(55000)
			cells("B4").PutValue(30000)
			cells("B5").PutValue(40000)
			cells("B6").PutValue(35000)
			cells("B7").PutValue(32000)
			cells("B8").PutValue(10000)
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'Initialize Style1
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

			'Set Font IsBold property
			style1.Font.IsBold = True

			'Set Style Alignment
			style1.HorizontalAlignment = TextAlignmentType.Center


			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set Style for A1 and B1
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)

			'Initialize Style2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set Font to Bold
			style2.Font.IsBold = False

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Style Pattern
			style2.Pattern = BackgroundType.Solid

			'Set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Loop over the cells
			For i As Integer = 1 To 7
				If i Mod 2 <> 0 Then
					'Apply Style
					cells(i, 0).SetStyle(style2)
				End If
			Next i

			'Initialize Style
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy the Style from another
			style3.Copy(style2)

			'Set cell format
			style3.Custom = """$""#,##0"


			'loop over the cells and Set Style
			For i As Integer = 1 To 7
				If i Mod 2 <> 0 Then
					cells(i, 1).SetStyle(style3)
				End If
			Next i

			'Initialize Style4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy Style from another
			style4.Copy(style2)

			'Set foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set Style pattern
			style4.Pattern = BackgroundType.Solid


			'Loop over the cells and set Style
			For i As Integer = 1 To 7
				If i Mod 2 = 0 Then
					cells(i, 0).SetStyle(style4)
				End If
			Next i

			'Initialize Style
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy the Style from Another
			style5.Copy(style4)

			'Set cell format
			style5.Custom = """$""#,##0"

			'Loop over the cells and set Style
			For i As Integer = 1 To 7
				If i Mod 2 = 0 Then
					cells(i, 1).SetStyle(style5)
				End If
			Next i
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'get index of newly added Worksheet
			Dim sheetIndex As Integer = workbook.Worksheets.Add()

			'Initialize Worksheet on given index
			Dim sheet As Worksheet = workbook.Worksheets(sheetIndex)

			'Set the name of worksheet
			sheet.Name = "Chart"

			'Create chart depending on selection made on Pie3DExploded
			Dim chartIndex As Integer = 0
			If CheckShow3D.Checked Then
				chartIndex = sheet.Charts.Add(ChartType.Pie3DExploded,1,1,25,10)
			Else
				chartIndex = sheet.Charts.Add(ChartType.PieExploded,1,1,25,10)
			End If

			'Initialize Chart
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set properties of chart like ForegroundColor
			chart.PlotArea.Area.ForegroundColor = Color.Coral

			'Set properties of chart like Boder Visibility
			chart.PlotArea.Border.IsVisible = False

			'Set properties of chart title
			chart.Title.Text = "Sales By Region"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set properties of nseries
			chart.NSeries.Add("Data!B2:B8", True)
			chart.NSeries.CategoryData = "Data!A2:A8"
			chart.NSeries.IsColorVaried = True

			'Set the legend position to Top
			chart.Legend.Position = LegendPositionType.Right
		End Sub
	End Class
End Namespace
