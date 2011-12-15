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
	''' Summary description for ClusteredColumn.
	''' </summary>
	Public Class ClusteredColumn
		Inherits System.Web.UI.Page
		Protected CategoryAxisTitle As System.Web.UI.WebControls.TextBox
		Protected ValueAxisTitle As System.Web.UI.WebControls.TextBox
		Protected ValueMaxValue As System.Web.UI.WebControls.DropDownList
		Protected ValueMinValue As System.Web.UI.WebControls.DropDownList
		Protected ValueMajorUnit As System.Web.UI.WebControls.DropDownList
		Protected ValueMinorUnit As System.Web.UI.WebControls.DropDownList
		Protected GapWidth As System.Web.UI.WebControls.DropDownList
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
			workbook.Save(HttpContext.Current.Response, "ClusteredColumn." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			Dim cells As Cells = workbook.Worksheets(0).Cells
			'Put string values for cells in first column
			cells("A1").PutValue("Region")
			cells("A2").PutValue("France")
			cells("A3").PutValue("Germany")
			cells("A4").PutValue("England")

			'Put Number values for cells in second column
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

			'Set Font property IsBold to Fase
			style1.Font.IsBold = True

			'Set Style Alignments
			style1.HorizontalAlignment = TextAlignmentType.Center
			style1.VerticalAlignment = TextAlignmentType.Center

			'Initalize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Set the width of the specified column
			cells.SetColumnWidth(1,15)
			cells.SetColumnWidth(1,15)

			'Apply style on cells A1 and B1
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)

			'Initialize Style2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set Font Property IsBold to False
			style2.Font.IsBold = False

			'Set Style Alignmet
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Style Patern
			style2.Pattern = BackgroundType.Solid

			'Apply Style of A2 and A4
			cells("A2").SetStyle(style2)
			cells("A4").SetStyle(style2)

			'Initialize Style3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy the properties from Style2
			style3.Copy(style2)

			'Set cell format
			style3.Custom = """$""#,##0"

			'Apply Style to cells B2 and B4 
			cells("B2").SetStyle(style3)
			cells("B4").SetStyle(style3)

			'Initialize Style4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy the properties of Style2
			style4.Copy(style2)

			'Sets foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set Style Pattern
			style4.Pattern = BackgroundType.Solid

			'Apply Style to cell A3
			cells("A3").SetStyle(style4)

			'Initialize Style5
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy properties from Style4
			style5.Copy(style4)

			'Set cell format
			style5.Custom = """$""#,##0"

			'Apply Stype to B3
			cells("B3").SetStyle(style5)
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'Initialize Worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the name of the worksheet
			sheet.Name = "Clustered Column"

			'Set Gridlines to Invisible
			sheet.IsGridlinesVisible = False

			'Create chart at next index on worksheet's chart collection
			Dim chartIndex As Integer = sheet.Charts.Add(ChartType.Column, 5, 1, 29, 10)

			'Initialize Chart
			Dim chart As Chart = sheet.Charts(chartIndex)

			'Add the nseries collection to a chart 
			chart.NSeries.Add("B2:B4", True)

			'Get or set the range of category axis values
			chart.NSeries.CategoryData = "A2:A4"
			chart.NSeries.IsColorVaried = True

			'Loop over the NSeries and Set DataLabels to Show Value
			For i As Integer = 0 To chart.NSeries.Count - 1
				chart.NSeries(i).DataLabels.ShowValue = True
			Next i

			'Set the legend position to Top
			chart.Legend.Position = LegendPositionType.Top
			chart.GapWidth = Integer.Parse(GapWidth.SelectedItem.Text)
			chart.CategoryAxis.MajorGridLines.IsVisible = False

			'Set properties of chart title
			chart.Title.Text = "Marketing Costs by Region"
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.Size = 12

			'Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = CategoryAxisTitle.Text
			chart.CategoryAxis.Title.TextFont.Color = Color.Black
			chart.CategoryAxis.Title.TextFont.IsBold = True
			chart.CategoryAxis.Title.TextFont.Size = 10

			'Set properties of valueaxis title
			chart.ValueAxis.Title.Text = ValueAxisTitle.Text
			chart.ValueAxis.Title.TextFont.Name = "Arial"
			chart.ValueAxis.Title.TextFont.Color = Color.Black
			chart.ValueAxis.Title.TextFont.IsBold = True
			chart.ValueAxis.Title.TextFont.Size = 10
			chart.ValueAxis.Title.RotationAngle = 90
			chart.ValueAxis.MajorUnit = Double.Parse(ValueMajorUnit.SelectedItem.Text)
			chart.ValueAxis.MaxValue = Double.Parse(ValueMaxValue.SelectedItem.Text)
			chart.ValueAxis.MinorUnit = Double.Parse(ValueMinorUnit.SelectedItem.Text)
			chart.ValueAxis.MinValue = Double.Parse(ValueMinValue.SelectedItem.Text)
		End Sub

	End Class
End Namespace
