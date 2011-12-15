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


Namespace Aspose.Cells.Demos.ChartTypes._3DCharts
	''' <summary>
	''' Summary description for Column3D.
	''' </summary>
	Public Class Column3D
		Inherits System.Web.UI.Page
		Protected WithEvents btnProcess As System.Web.UI.WebControls.Button
		Protected ChartTypeList As System.Web.UI.WebControls.DropDownList
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
			workbook.Save(HttpContext.Current.Response, "Column3D." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			Dim sheet As Worksheet = workbook.Worksheets(0)
			sheet.IsGridlinesVisible = False

			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put values in row cells of Column 1
			cells("A1").PutValue("Region")
			cells("B1").PutValue("Apple")
			cells("C1").PutValue("Orange")

			'Put values in row cells of Column 2
			cells("A2").PutValue("France")
			cells("B2").PutValue(800000)
			cells("C2").PutValue(300000)

			'Put values in row cells of Column 3
			cells("A3").PutValue("Germany")
			cells("B3").PutValue(200000)
			cells("C3").PutValue(600000)

			'Put values in row cells of Column 4
			cells("A4").PutValue("England")
			cells("B4").PutValue(400000)
			cells("C4").PutValue(600000)
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'initialize Style1
			Dim style1 As Style = workbook.Styles(workbook.Styles.Add())
			'Set border Style
			style1.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'Set Font IsBold Property to True for Header
			style1.Font.IsBold = True

			'Set Style Alignment
			style1.HorizontalAlignment = TextAlignmentType.Center
			style1.VerticalAlignment = TextAlignmentType.Center

			'initialize Cells of Worksheet[0]
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set Style of Column Headers
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)
			cells("C1").SetStyle(style1)

			'Initialize Style 2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set FOnt IsBold property to False
			style2.Font.IsBold = False

			'Set Style Alignmment
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Style Patern
			style2.Pattern = BackgroundType.Solid

			'Apply Style on Cells A2 and A4
			cells("A2").SetStyle(style2)
			cells("A4").SetStyle(style2)

			'Initialize Style 3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style3.Copy(style2)

			'Set cell format
			style3.Custom = """$""#,##0"

			'Apply Style on Cells B2, C2, B4 and C4
			cells("B2").SetStyle(style3)
			cells("C2").SetStyle(style3)
			cells("B4").SetStyle(style3)
			cells("C4").SetStyle(style3)

			'Initialize Style 4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style4.Copy(style3)

			'Set foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			style4.Pattern = BackgroundType.Solid

			'Apply Style on Cells A3
			cells("A3").SetStyle(style4)

			'Initialize Style 5
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style5.Copy(style4)

			'Set cell format
			style5.Custom = """$""#,##0"

			'Set Style of cells B3 and C3
			cells("B3").SetStyle(style5)
			cells("C3").SetStyle(style5)
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'Get index of newly added worksheet
			Dim sheetIndex As Integer = workbook.Worksheets.Add(SheetType.Chart)

			'initialize Worksheet
			Dim sheet As Worksheet = workbook.Worksheets(sheetIndex)

			'Set the name of worksheet
			sheet.Name = "Column3D Chart"

			'Create Chart depending on selected value on ChartTypeList
			Dim indexChart As Integer = 0
			Select Case ChartTypeList.SelectedItem.Text
				Case "Cylinder"
					indexChart = sheet.Charts.Add(ChartType.CylindricalColumn3D, 0, 0, 0, 0)
				Case "Cone"
					indexChart = sheet.Charts.Add(ChartType.ConicalColumn3D,0,0,0,0)
				Case "Pyramid"
					indexChart = sheet.Charts.Add(ChartType.PyramidColumn3D,0,0,0,0)
			End Select

			Dim chart As Chart = sheet.Charts(indexChart)
			chart.PlotArea.Border.IsVisible = False

			'Set Properties of chart title
			chart.Title.Text = "Fruit Sales By Region"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set properties of nseries
			chart.NSeries.Add("Sheet1!B2:C4",True)

			'Set nseries Category Data source
			chart.NSeries.CategoryData = "Sheet1!A2:A4"

			'Set nseries Color varience to True
			chart.NSeries.IsColorVaried = True

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Loop on Nseriese and Name them as values in cells
			For i As Integer = 0 To chart.NSeries.Count - 1
				chart.NSeries(i).Name = cells(0,i+1).Value.ToString()
			Next i

			'Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Region"
			chart.CategoryAxis.Title.TextFont.Color = Color.Black
			chart.CategoryAxis.Title.TextFont.IsBold = True
			chart.CategoryAxis.Title.TextFont.Size = 10

			'Set properties of legend
			chart.Legend.Position = LegendPositionType.Top
		End Sub
	End Class
End Namespace
