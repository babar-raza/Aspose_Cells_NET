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
	''' Summary description for Radar.
	''' </summary>
	Public Class Radar
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
			Me.ID = "Radar"
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
			workbook.Save(HttpContext.Current.Response, "Radar." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Set the name of worksheet
			sheet.Name = "Data"
			sheet.IsGridlinesVisible = False

			Dim cells As Cells = workbook.Worksheets(0).Cells
			'Put values fo row 1 (column Header)
			cells("A1").PutValue("Brand Name")
			cells("B1").PutValue("Vitamin A")
			cells("C1").PutValue("Vitamin B1")
			cells("D1").PutValue("Vitamin B2")
			cells("E1").PutValue("Vitamin C")
			cells("F1").PutValue("Vitamin D")
			cells("G1").PutValue("Vitamin E")

			'Put Values for row 2
			cells("A2").PutValue("Brand A")
			cells("B2").PutValue(100)
			cells("C2").PutValue(100)
			cells("D2").PutValue(100)
			cells("E2").PutValue(80)
			cells("F2").PutValue(100)
			cells("G2").PutValue(70)

			'put values for row 3
			cells("A3").PutValue("Brand B")
			cells("B3").PutValue(80)
			cells("C3").PutValue(75)
			cells("D3").PutValue(80)
			cells("E3").PutValue(100)
			cells("F3").PutValue(50)
			cells("G3").PutValue(15)

			'put values for row 4
			cells("A4").PutValue("Brand C")
			cells("B4").PutValue(40)
			cells("C4").PutValue(25)
			cells("D4").PutValue(40)
			cells("E4").PutValue(55)
			cells("F4").PutValue(30)
			cells("G4").PutValue(10)
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'Initialize Style1
			Dim style1 As Style = workbook.Styles(workbook.Styles.Add())
			'Set border for Style
			style1.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'Set Style Font to Bold
			style1.Font.IsBold = True

			'Set Style Alignment
			style1.HorizontalAlignment = TextAlignmentType.Center


			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set column Width
			cells.SetColumnWidth(0, 12)

			'loop over the cells 
			For i As Integer = 1 To 6
				'Set the width of the specified column 
				cells.SetColumnWidth(i, 9)
			Next i

			'Set style of column header
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)
			cells("C1").SetStyle(style1)
			cells("D1").SetStyle(style1)
			cells("E1").SetStyle(style1)
			cells("F1").SetStyle(style1)
			cells("G1").SetStyle(style1)

			'initialize Style 2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'set style font to Normal
			style2.Font.IsBold = False

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Style pattern
			style2.Pattern = BackgroundType.Solid

			'Set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right

			'loop over the cells and Set style
			For i As Integer = 1 To 3
				If i Mod 2 <> 0 Then
					cells(i, 0).SetStyle(style2)
					cells(i, 1).SetStyle(style2)
					cells(i, 2).SetStyle(style2)
					cells(i, 3).SetStyle(style2)
					cells(i, 4).SetStyle(style2)
					cells(i, 5).SetStyle(style2)
					cells(i, 6).SetStyle(style2)
				End If
			Next i

			'initialize Style 
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style3.Copy(style2)

			'Set Style ForegroundColor
			style3.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set Style Pattern
			style3.Pattern = BackgroundType.Solid

			'Loop over the cells and Set Style
			For i As Integer = 1 To 3
				If i Mod 2 = 0 Then
					cells(i, 0).SetStyle(style2)
					cells(i, 1).SetStyle(style3)
					cells(i, 2).SetStyle(style3)
					cells(i, 3).SetStyle(style3)
					cells(i, 4).SetStyle(style3)
					cells(i, 5).SetStyle(style3)
					cells(i, 6).SetStyle(style3)
				End If
			Next i
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'get index of newly added Worksheet
			Dim sheetIndex As Integer = workbook.Worksheets.Add()

			'initialize worksheet for given Index
			Dim sheet As Worksheet = workbook.Worksheets(sheetIndex)

			'Set the name of worksheet
			sheet.Name = "Chart"

			'Create chart depending on the ChartTypeList's SelectedItem
			Dim chartIndex As Integer = 0
			Select Case ChartTypeList.SelectedItem.Text
				Case "Radar"
					chartIndex = sheet.Charts.Add(ChartType.Radar,5,1,29,10)
				Case "RadarWithDataMarkers"
					chartIndex = sheet.Charts.Add(ChartType.RadarWithDataMarkers,5,1,29,10)
			End Select
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set properties of chart
			chart.PlotArea.Border.IsVisible = False

			'Set properties of chart title
			chart.Title.Text = "Nutritional Analysis"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set properties of nseries
			chart.NSeries.Add("B2:G4",False)
			chart.NSeries.CategoryData = "B1:G1"

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'loop over the NSeries
			For i As Integer = 0 To chart.NSeries.Count - 1
				'Set NSeries Name to values from cells
				chart.NSeries(i).Name = cells("A" & (i + 2).ToString()).Value.ToString()
			Next i
		End Sub

	End Class
End Namespace
