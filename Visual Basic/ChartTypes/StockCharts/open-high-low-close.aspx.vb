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
	''' Summary description for OpenHighLowClose.
	''' </summary>
	Public Class OpenHighLowClose
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
			workbook.Save(HttpContext.Current.Response, "OpenHighLowClose." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Initialize Worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the name of worksheet
			sheet.Name = "Data"

			'Set Gridlines invisible
			sheet.IsGridlinesVisible = False

			'Initilize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put values for Column Header
			cells("A1").PutValue("Company Name")
			cells("B1").PutValue("Open")
			cells("C1").PutValue("High")
			cells("D1").PutValue("Low")
			cells("E1").PutValue("Close")

			'Put values for Row 1
			cells("A2").PutValue("Microsoft")
			cells("B2").PutValue(21.00)
			cells("C2").PutValue(27.20)
			cells("D2").PutValue(23.49)
			cells("E2").PutValue(25.45)


			'Put values for Row 2
			cells("A3").PutValue("Mutual Fund 1")
			cells("B3").PutValue(28.52)
			cells("C3").PutValue(25.03)
			cells("D3").PutValue(19.55)
			cells("E3").PutValue(23.05)

			'Put values for Row 3
			cells("A4").PutValue("Mutual Fund 2")
			cells("B4").PutValue(9.05)
			cells("C4").PutValue(19.05)
			cells("D4").PutValue(15.12)
			cells("E4").PutValue(17.32)
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
			cells.SetColumnWidth(0, 15)

			'Set Style for Header
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)
			cells("C1").SetStyle(style1)
			cells("D1").SetStyle(style1)

			'Initialize Style 2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set Font to Normal
			style2.Font.IsBold = False

			'Set Style Alignment
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Style Pattern
			style2.Pattern = BackgroundType.Solid

			'Set Style
			cells("A2").SetStyle(style2)
			cells("A4").SetStyle(style2)

			'Initialize Style 3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style3.Copy(style2)

			'Set cell format
			style3.Number = 2


			'Loop over the cells and Set Style
			For i As Integer = 1 To 3
				If i Mod 2 <> 0 Then
					cells(i, 1).SetStyle(style3)
					cells(i, 2).SetStyle(style3)
					cells(i, 3).SetStyle(style3)
				End If
			Next i

			'Initialize Style 4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy the properties from another Style Object
			style4.Copy(style2)

			'Set foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set Style Pattern
			style4.Pattern = BackgroundType.Solid

			'Apply Style
			cells("A3").SetStyle(style4)

			'Initialize Style 4
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy the properties from another Style Object
			style5.Copy(style4)

			'Set cell format
			style5.Number = 2

			'Loop over cells and Set Style
			For i As Integer = 1 To 3
				If i Mod 2 = 0 Then
					cells(i, 1).SetStyle(style5)
					cells(i, 2).SetStyle(style5)
					cells(i, 3).SetStyle(style5)
				End If
			Next i
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'Get index of newly added Worksheet
			Dim sheetIndex As Integer = workbook.Worksheets.Add()

			'Initialize Worksheet for given index
			Dim sheet As Worksheet = workbook.Worksheets(sheetIndex)

			'Set the name of worksheet
			sheet.Name = "Chart"

			'Create chart of Type 	StockOpenHighLowClose
			Dim chartIndex As Integer = sheet.Charts.Add(ChartType.StockOpenHighLowClose,1,1,25,10)

			'Initialize Chart
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set pproperties of nseries
			chart.NSeries.Add("Data!B2:E4",True)
			chart.NSeries.CategoryData = "Data!A2:A4"

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'loop over NSeries
			For i As Integer = 0 To chart.NSeries.Count - 1
				'Set Name from values of cells
				chart.NSeries(i).Name = cells(0,i+1).Value.ToString()
			Next i

			'Set properties of chart title
			chart.Title.Text = " Stock chart"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set Properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Scock Names"
			chart.CategoryAxis.Title.TextFont.Color = Color.Black
			chart.CategoryAxis.Title.TextFont.Size = 10
			chart.CategoryAxis.Title.TextFont.IsBold = True

			'Set properties of valueaxis title
			chart.ValueAxis.Title.Text= "Stock Price"
			chart.ValueAxis.Title.TextFont.Color = Color.Black
			chart.ValueAxis.Title.TextFont.IsBold = True
			chart.ValueAxis.Title.TextFont.Size =10
			chart.ValueAxis.Title.RotationAngle = 90
		End Sub
	End Class
End Namespace
