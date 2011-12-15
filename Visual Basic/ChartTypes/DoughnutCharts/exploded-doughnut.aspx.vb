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
	''' Summary description for ExplodedDoughnut.
	''' </summary>
	Public Class ExplodedDoughnut
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
			workbook.Save(HttpContext.Current.Response, "ExplodedDoughnut." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Initialize Worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the name of worksheet
			sheet.Name = "Data"

			'Set Gridlines of worksheet invisible
			sheet.IsGridlinesVisible = False

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put values into cells
			cells("A1").PutValue("Product Name")
			cells("A2").PutValue("Apple")
			cells("A3").PutValue("Orange")

			cells("B1").PutValue(2006)

			cells("B2").PutValue(30000)
			cells("B3").PutValue(50000)
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
			style1.VerticalAlignment = TextAlignmentType.Center

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set the width of the specified column 
			cells.SetColumnWidth(0, 15)

			'Apply Style on B1 and A1 (first row)
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)

			'Initialize Style2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set IsBold property of Font to False
			style2.Font.IsBold = False

			'Set Alignment of Style
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Pattern of Sty;e
			style2.Pattern = BackgroundType.Solid

			'Apply style to A2
			cells("A2").SetStyle(style2)

			'Initialize Style3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style3.Copy(style2)

			'Set cell format
			style3.Custom = """$""#,##0"

			'Apply Style to B2
			cells("B2").SetStyle(style3)

			'Initialize Style4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style4.Copy(style2)

			'Set foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set Style Pattern
			style4.Pattern = BackgroundType.Solid

			'Apply Style to A3
			cells("A3").SetStyle(style4)

			'Initialize Style5
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style5.Copy(style4)


			'Set cell format
			style5.Custom = """$""#,##0"

			'Apply Style to B3
			cells("B3").SetStyle(style5)
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'Get index of newly added Worksheet
			Dim sheetIndex As Integer = workbook.Worksheets.Add()

			'Initialize Worksheet of given index
			Dim sheet As Worksheet = workbook.Worksheets(sheetIndex)

			'Set the name of worksheet
			sheet.Name = "Chart"

			'Create chart of type DoughnutExploded
			Dim chartIndex As Integer = sheet.Charts.Add(ChartType.DoughnutExploded,1,1,27,10)
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Set the properties of chart to set border invisible
			chart.PlotArea.Border.IsVisible = False

			'Set the properties of chart title
			chart.Title.Text = "Fruit Sales by Region For Years"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set the properties of nseries
			chart.NSeries.Add("Data!B2:B3",True)

			'Set Nseries Category Datasource
			chart.NSeries.CategoryData = "Data!A2:A3"
			chart.NSeries.IsColorVaried = True

			'Set the properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Region"
			chart.CategoryAxis.Title.TextFont.Color = Color.Black
			chart.CategoryAxis.Title.TextFont.IsBold = True
			chart.CategoryAxis.Title.TextFont.Size = 10

			'Set the properties of valueaxis title
			chart.ValueAxis.Title.Text = "Thousand"
			chart.ValueAxis.Title.TextFont.Color = Color.Black
			chart.ValueAxis.Title.TextFont.IsBold = True
			chart.ValueAxis.Title.TextFont.Size = 10

			'Set the properties of legend to show on Top
			chart.Legend.Position = LegendPositionType.Top
		End Sub
	End Class
End Namespace
