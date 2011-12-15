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
	''' Summary description for Surface3D.
	''' </summary>
	Public Class Surface3D
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
			workbook.Save(HttpContext.Current.Response, "3DSurface." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			Dim cells As Cells = workbook.Worksheets(0).Cells
			'Put a value into a cell
			cells("A1").PutValue("Temperature")
			cells("A2").PutValue("Seconds")
			cells("A3").PutValue(0.2)
			cells("A4").PutValue(0.3)
			cells("A5").PutValue(0.4)
			cells("A6").PutValue(0.5)
			cells("A7").PutValue(0.6)
			cells("A8").PutValue(0.7)
			cells("A9").PutValue(0.8)
			cells("A10").PutValue(0.9)
			cells("A11").PutValue(1)

			'Merge a specified range of cells into a single cell
			cells.Merge(0,1,2,1)
			cells("B1").PutValue(10)
			cells("B3").PutValue(99)
			cells("B4").PutValue(107)
			cells("B5").PutValue(119)
			cells("B6").PutValue(135)
			cells("B7").PutValue(155)
			cells("B8").PutValue(184)
			cells("B9").PutValue(193)
			cells("B10").PutValue(295)
			cells("B11").PutValue(384)

			'Merge a specified range of cells into a single cell
			cells.Merge(0,2,2,1)
			cells("C1").PutValue(20)
			cells("C3").PutValue(175)
			cells("C4").PutValue(185)
			cells("C5").PutValue(200)
			cells("C6").PutValue(220)
			cells("C7").PutValue(245)
			cells("C8").PutValue(279)
			cells("C9").PutValue(349)
			cells("C10").PutValue(385)
			cells("C11").PutValue(499)

			'Merge a specified range of cells into a single cell
			cells.Merge(0,3,2,1)
			cells("D1").PutValue(30)
			cells("D3").PutValue(250)
			cells("D4").PutValue(260)
			cells("D5").PutValue(275)
			cells("D6").PutValue(275)
			cells("D7").PutValue(320)
			cells("D8").PutValue(356)
			cells("D9").PutValue(392)
			cells("D10").PutValue(405)
			cells("D11").PutValue(459)

			'Merge a specified range of cells into a single cell
			cells.Merge(0,4,2,1)
			cells("E1").PutValue(40)
			cells("E3").PutValue(467)
			cells("E4").PutValue(385)
			cells("E5").PutValue(349)
			cells("E6").PutValue(279)
			cells("E7").PutValue(245)
			cells("E8").PutValue(220)
			cells("E9").PutValue(200)
			cells("E10").PutValue(185)
			cells("E11").PutValue(175)

			'Merge a specified range of cells into a single cell
			cells.Merge(0,5,2,1)
			cells("F1").PutValue(50)
			cells("F3").PutValue(400)
			cells("F4").PutValue(305)
			cells("F5").PutValue(209)
			cells("F6").PutValue(192)
			cells("F7").PutValue(163)
			cells("F8").PutValue(144)
			cells("F9").PutValue(118)
			cells("F10").PutValue(59)
			cells("F11").PutValue(25)
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'Initialize Style
			Dim style1 As Style = workbook.Styles(workbook.Styles.Add())
			'Set border setting of Style
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

			'Set alignmet of Style
			style1.HorizontalAlignment = TextAlignmentType.Center

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set the width of the specified column 
			cells.SetColumnWidth(0, 12)

			'Set Style for two rows
			cells("A1").SetStyle(style1)
			cells("A2").SetStyle(style1)
			cells("B1").SetStyle(style1)
			cells("B2").SetStyle(style1)
			cells("C1").SetStyle(style1)
			cells("C2").SetStyle(style1)
			cells("D1").SetStyle(style1)
			cells("D2").SetStyle(style1)
			cells("E1").SetStyle(style1)
			cells("E2").SetStyle(style1)
			cells("F1").SetStyle(style1)
			cells("F2").SetStyle(style1)

			'initialize Style2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set Font to Normal
			style2.Font.IsBold = False

			'Set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)
			style2.Pattern = BackgroundType.Solid

			'Loop over Cells and apply Style
			For i As Integer = 2 To 10
				If i Mod 2 = 0 Then
					cells(i, 0).SetStyle(style2)
					cells(i, 1).SetStyle(style2)
					cells(i, 2).SetStyle(style2)
					cells(i, 3).SetStyle(style2)
					cells(i, 4).SetStyle(style2)
					cells(i, 5).SetStyle(style2)
				End If
			Next i
			'initialize Style3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy Style properties from another Object
			style3.Copy(style2)

			'Set foreground color
			style3.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set Pattern of Style
			style3.Pattern = BackgroundType.Solid

			'Loop over the cells and set style
			For i As Integer = 2 To 10
				If i Mod 2 <> 0 Then
					cells(i, 0).SetStyle(style2)
					cells(i, 1).SetStyle(style3)
					cells(i, 2).SetStyle(style3)
					cells(i, 3).SetStyle(style3)
					cells(i, 4).SetStyle(style3)
					cells(i, 5).SetStyle(style3)
				End If
			Next i
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Set the name of worksheet
			sheet.Name = "3D Surface"
			sheet.IsGridlinesVisible = False

			'Create chart 
			Dim chartIndex As Integer = 0
			Select Case ChartTypeList.SelectedItem.Text
				Case "Surface3D"
					chartIndex = sheet.Charts.Add(ChartType.Surface3D,1,7,25,16)
				Case "SurfaceWireframe3D"
					chartIndex = sheet.Charts.Add(ChartType.SurfaceWireframe3D,1,7,25,16)
			End Select
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set properties of chart
			chart.PlotArea.Border.IsVisible = False

			'Set properties of nseries
			chart.NSeries.Add("B3:F11",True)
			chart.NSeries.CategoryData = "A3:A11"
			chart.NSeries.IsColorVaried = True

			Dim cells As Cells = workbook.Worksheets(0).Cells

			For i As Integer = 0 To chart.NSeries.Count - 1
				chart.NSeries(i).Name = cells(0, i + 1).StringValue
			Next i

			'Set properties of chart title
			chart.Title.Text = "Tensile strenth Measurements"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Seconds"
			chart.CategoryAxis.Title.TextFont.Color = Color.Black
			chart.CategoryAxis.Title.TextFont.IsBold = True
			chart.CategoryAxis.Title.TextFont.Size = 10

			'Set properties of valueaxis title
			chart.ValueAxis.Title.Text = "Tensile Strength"
			chart.ValueAxis.Title.TextFont.Color = Color.Black
			chart.ValueAxis.Title.TextFont.IsBold = True
			chart.ValueAxis.Title.TextFont.Size = 10
			chart.ValueAxis.Title.RotationAngle = 90

		End Sub
	End Class
End Namespace
