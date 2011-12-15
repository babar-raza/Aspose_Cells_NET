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
	''' Summary description for PercentStackedBar.
	''' </summary>
	Public Class PercentStackedBar
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
			workbook.Save(HttpContext.Current.Response, "PercentStackedBar." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Initialize worksheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Set the name of worksheet
			worksheet.Name = "Data"

			'Set GridLines Invisible
			worksheet.IsGridlinesVisible = False

			'initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'put values for row 1
			cells("A1").PutValue("Product Name")
			cells("B1").PutValue("Quarter1")
			cells("C1").PutValue("Quarter2")
			cells("D1").PutValue("Quarter3")
			cells("E1").PutValue("Quarter4")

			'put values for row 2
			cells("A2").PutValue("Product1")
			cells("B2").PutValue(0.33)
			cells("C2").PutValue(0.21)
			cells("D2").PutValue(0.35)
			cells("E2").PutValue(0.22)

			'put values for row 3
			cells("A3").PutValue("Product2")
			cells("B3").PutValue(0.17)
			cells("C3").PutValue(0.54)
			cells("D3").PutValue(0.17)
			cells("E3").PutValue(0.60)

			'put values for row 4
			cells("A4").PutValue("Product3")
			cells("B4").PutValue(0.50)
			cells("C4").PutValue(0.25)
			cells("D4").PutValue(0.48)
			cells("E4").PutValue(0.18)
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'Initialize Style1
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

			'Set Styles alignment
			style1.HorizontalAlignment = TextAlignmentType.Center
			style1.VerticalAlignment = TextAlignmentType.Center

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set the width of the specified column 
			cells.SetColumnWidth(0,12)

			'Set style of Row 1
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)
			cells("C1").SetStyle(style1)
			cells("D1").SetStyle(style1)
			cells("E1").SetStyle(style1)

			'initialize Style2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set Font property IsBold to false
			style2.Font.IsBold = False

			'set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Pattern of Style
			style2.Pattern = BackgroundType.Solid

			'Set Style to cell A2 and A4
			cells("A2").SetStyle(style2)
			cells("A4").SetStyle(style2)

			'Initialize Style3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy properties from Style2
			style3.Copy(style2)

			'Set cell format
			style3.Number = 9

			'Loop over cells and Appky Style
			For i As Integer = 1 To 3
				If i Mod 2 <>0 Then
					cells(i, 1).SetStyle(style3)
					cells(i,2).SetStyle(style3)
					cells(i,3).SetStyle(style3)
					cells(i,4).SetStyle(style3)
				End If
			Next i

			'Initialize Style4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy properties from Style2
			style4.Copy(style2)

			'Sets foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)
			style4.Pattern = BackgroundType.Solid

			'Set Style of cell A3
			cells("A3").SetStyle(style4)

			'Initialize Style5
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy Style properties from Style4
			style5.Copy(style4)

			'Set cell format
			style5.Number = 9

			'Loop over cells and Appky Style
			For i As Integer = 1 To 3
				If i Mod 2 = 0 Then
					cells(i,1).SetStyle(style5)
					cells(i,2).SetStyle(style5)
					cells(i,3).SetStyle(style5)
					cells(i,4).SetStyle(style5)
				End If
			Next i
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'Initialize worksheet of type Chart
			Dim sheetIndex As Integer = workbook.Worksheets.Add(SheetType.Chart)
			Dim sheet As Worksheet = workbook.Worksheets(sheetIndex)

			'Set the name of worksheet
			sheet.Name = "Chart"

			'Create chart, If Check box on Ui is Checked then create Bar3D100PercentStacked Chart else Bar100PercentStacked Chart
			Dim chartIndex As Integer = 0
			If CheckBoxShow3D.Checked Then
				chartIndex = sheet.Charts.Add(ChartType.Bar3D100PercentStacked,0,0,0,0)
			Else
				chartIndex = sheet.Charts.Add(ChartType.Bar100PercentStacked,0,0,0,0)
			End If
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set properties of chart and hide gridLines based upon the state od check box on UI
			chart.CategoryAxis.MajorGridLines.IsVisible = False
			If CheckBoxShow3D.Checked Then
			  chart.PlotArea.Border.IsVisible = False
			End If

			'Set properties of chart title
			chart.Title.Text = "Product contribution to total sales"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set properties of nseries
			chart.NSeries.Add("Data!B2:E4", False)

			'Set Category Data from cells between B1 and E1
			chart.NSeries.CategoryData = "Data!B1:E1"

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'loop over the Chart's Nseries and Assign Name from Cell Values
			For i As Integer = 0 To chart.NSeries.Count - 1
				chart.NSeries(i).Name = cells(i+1,0).Value.ToString()
			Next i

			'Set properties of valueaxis title
			chart.ValueAxis.Title.Text = "% of total sales"
			chart.ValueAxis.Title.TextFont.Color = Color.Black
			chart.ValueAxis.Title.TextFont.IsBold = True
			chart.ValueAxis.Title.TextFont.Size = 10
		End Sub



	End Class
End Namespace
