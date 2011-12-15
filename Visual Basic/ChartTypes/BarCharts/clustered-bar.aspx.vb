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
	''' Summary description for ClusteredBar.
	''' </summary>
	Public Class ClusteredBar
		Inherits System.Web.UI.Page
		Protected checkBoxShow3D As System.Web.UI.WebControls.CheckBox
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

		''' <summary>
		''' Initialize Workbook, insert dummy data
		''' Create chart based on dummy data
		''' save the file in format selected from UI
		''' </summary>
		''' <param name="sender"></param>
		''' <param name="e"></param>
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
			workbook.Save(HttpContext.Current.Response, "ClusteredBar." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))
			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Put values in rows of cloumn A
			cells("A1").PutValue("Region")
			cells("A2").PutValue("France")
			cells("A3").PutValue("Germany")
			cells("A4").PutValue("England")

			'Put values into a cell B1 and C1
			cells("B1").PutValue("Apple")
			cells("C1").PutValue("Orange")

			'Put number type values in row 2, 3, 4 for Column B
			cells("B2").PutValue(220000)
			cells("B3").PutValue(80000)
			cells("B4").PutValue(150000)

			'Put number type values in row 2, 3, 4 for Column C
			cells("C2").PutValue(100000)
			cells("C3").PutValue(150000)
			cells("C4").PutValue(60000)
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'Initialize Style1
			Dim style1 As Style = workbook.Styles(workbook.Styles.Add())

			'Set border settings for style1
			style1.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style1.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'Set Font property IsBold to False
			style1.Font.IsBold = True

			'Set alignment for Style1
			style1.HorizontalAlignment = TextAlignmentType.Center

			'intialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Apply Style1 on A1, B1,C1
			cells("A1").SetStyle(style1)
			cells("B1").SetStyle(style1)
			cells("C1").SetStyle(style1)


			'Intialize style2
			Dim style2 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy data from another style object
			style2.Copy(style1)

			'Set isBold property of style Font to False
			style2.Font.IsBold = False

			'Set alignment of style2
			style2.HorizontalAlignment = TextAlignmentType.Right

			'Set foreground color
			style2.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)

			'Set Pattern of Style2
			style2.Pattern = BackgroundType.Solid

			'Apply style2 to A2 and A4
			cells("A2").SetStyle(style2)
			cells("A4").SetStyle(style2)

			'Initialize Style3
			Dim style3 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy properties from Style2
			style3.Copy(style2)

			'Set cell format
			style3.Custom = """$""#,##0"

			'Apply Style to B2, C2, B4 and C4
			cells("B2").SetStyle(style3)
			cells("C2").SetStyle(style3)
			cells("B4").SetStyle(style3)
			cells("C4").SetStyle(style3)

			'Initialize Style4
			Dim style4 As Style = workbook.Styles(workbook.Styles.Add())

			'copy contents from Style2
			style4.Copy(style2)

			'Sets foreground color
			style4.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)

			'Set style Pattern to solid
			style4.Pattern = BackgroundType.Solid

			'Apply Style to A3
			cells("A3").SetStyle(style4)

			'Initialize Style5
			Dim style5 As Style = workbook.Styles(workbook.Styles.Add())

			'Copy properties from Style4
			style5.Copy(style4)

			'Set cell format
			style5.Custom = """$""#,##0"

			'Set Style to B3 and C3
			cells("B3").SetStyle(style5)
			cells("C3").SetStyle(style5)
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'Initialize worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the name of worksheet
			sheet.Name = "Clustered Bar"

			'Set GridLines invisible
			sheet.IsGridlinesVisible = False

			'Create chart, If Check box on Ui is Checked then create Bar3DClustered Chart else Bar Chart
			Dim indexChart As Integer = 0
			If checkBoxShow3D.Checked Then
				indexChart = sheet.Charts.Add(ChartType.Bar3DClustered,5,1,26,10)
			Else
				indexChart = sheet.Charts.Add(ChartType.Bar,5,1,26,10)
			End If
			Dim chart As Chart = sheet.Charts(indexChart)

			'Set properties of chart based upon state of check box on UI
			If checkBoxShow3D.Checked Then
				chart.PlotArea.Border.IsVisible = False
			End If
			chart.CategoryAxis.MajorGridLines.IsVisible = False

			'Set properties of chart title
			chart.Title.Text = "Fruit Sales By Region"
			chart.Title.TextFont.Color = Color.Black
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 12

			'Set properties of nseries
			chart.NSeries.Add("B2:C4", True)
			chart.NSeries.CategoryData = "A2:A4"
			chart.NSeries.IsColorVaried = True

			'Initialize Cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'loop over the Chart's Nseries and Assign Name from Cell Values
			For i As Integer = 0 To chart.NSeries.Count - 1
				chart.NSeries(i).Name = cells(0,i+1).Value.ToString()
			Next i

			'Set properties of categoryaxis title
			chart.CategoryAxis.Title.Text = "Region"
			chart.CategoryAxis.Title.TextFont.Color = Color.Black
			chart.CategoryAxis.Title.TextFont.IsBold = True
			chart.CategoryAxis.Title.TextFont.Size = 10
			chart.CategoryAxis.Title.RotationAngle = 90

			'Set properties of legend to show on Top
			chart.Legend.Position = LegendPositionType.Top
		End Sub



	End Class
End Namespace
