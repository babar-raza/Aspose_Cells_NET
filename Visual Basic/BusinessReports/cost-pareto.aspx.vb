Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Configuration
Imports System.Collections
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Drawing
Imports Aspose.Cells
Imports Aspose.Cells.Drawing
Imports Aspose.Cells.Charts


Namespace Aspose.Cells.Demos
	Partial Public Class CostPareto
		Inherits System.Web.UI.Page
		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub btnProcess_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create a dataset object
			Dim ds As New DataSet()

			'Get data from xml file
			Dim path As String = MapPath(".")
			path = path.Substring(0, path.LastIndexOf("\"))
			path &= "\Database\CostPareto.xml"

			'Load data from xml file to dataset
			ds.ReadXml(path, XmlReadMode.ReadSchema)

			'Create a new workbook
			Dim workbook As New Workbook()

			'Generate first data sheet
			GenerateDataSheet(workbook, ds)

			'Generate second chart sheet
			GenerateChartSheet(workbook, ds)

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
			workbook.Save(HttpContext.Current.Response, "CostPareto." & ddlFileVersion.SelectedItem.Value.ToLower(), ContentDisposition.Attachment, New XlsSaveOptions(saveFormat))

			' note by Vit - end response to avoid unneeded html after xls
			Response.End()
		End Sub


		Private Sub GenerateDataSheet(ByVal workbook As Workbook, ByVal ds As DataSet)
			'Write data to first data sheet
			Dim sheet1 As Worksheet = workbook.Worksheets(0)

			'Name the sheet
			sheet1.Name = "Cost Data"

			'Write sheet1 cells data to cells object
			Dim cells As Cells = sheet1.Cells

			'Import data into cells
			cells.ImportDataTable(ds.Tables(0), True, 0, 0, ds.Tables(0).Rows.Count, ds.Tables(0).Columns.Count)

			'Set header style with specific formatting attributes
			Dim styles As StyleCollection = workbook.Styles

			'Set style index
			Dim styleIndex As Integer = styles.Add()

			'Set style attribute using style index
			Dim style As Style = styles(styleIndex)

			'Set font size 
			style.Font.Size = 10

			'Set font color to white
			style.Font.Color = Color.White

			'Set font to bold
			style.Font.IsBold = True

			'Set font name to Verdana
			style.Font.Name = "Verdana"

			'Locked style
			style.IsLocked = True

			'Set vertical alignment 
			style.VerticalAlignment = TextAlignmentType.Center

			'Set horizontal alignment
			style.HorizontalAlignment = TextAlignmentType.Left

			'Set indent level
			style.IndentLevel = 1

			'Set top, bottom, left and right borders style
			style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thick
			style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'Change the palette for the spreadsheet in the specified index
			workbook.ChangePalette(Color.FromArgb(10, 100, 180), 50)

			'Change foreground color
			style.ForegroundColor = Color.FromArgb(10, 100, 180)

			'Set background style pattern
			style.Pattern = BackgroundType.Solid

			'Set first two column's widths and set the height of the first row
			cells.SetColumnWidth(0, 25)
			cells.SetColumnWidth(1, 18)
			cells.SetRowHeight(0, 30)

			'Apply the style to A1 cell
			cells(0, 0).SetStyle(style)

			'Add a new style
			styleIndex = styles.Add()
			Dim style1 As Style = styles(styleIndex)

			'Copy above created style to it
			style1.Copy(style)

			'Set horizontal alignment and indentation
			style1.HorizontalAlignment = TextAlignmentType.Right
			style1.IndentLevel = 0

			'Apply the style to B1 cell
			cells(0, 1).SetStyle(style1)

			'Set current row to 1
			Dim currentRow As Integer = 1
			For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
				'Set row height and color
				cells.SetRowHeight(currentRow, 20)
				Dim color As Color = Color.FromArgb(255, 255, 255)

				'Change palette color of workbook
				workbook.ChangePalette(color, 51)

				'Change color of even number rows
				If currentRow Mod 2 = 0 Then
					'Set color
					color = Color.FromArgb(250, 250, 200)

					'Change palette color of workbook
					workbook.ChangePalette(color, 52)
				End If

				'Set style for the first column cells
				styleIndex = styles.Add()

				'Set style attribute using style index
				Dim styleCell1 As Style = styles(styleIndex)

				'Set font size
				styleCell1.Font.Size = 10

				'Set font name 
				styleCell1.Font.Name = "Arial"

				'Set horizontal alignment
				styleCell1.HorizontalAlignment = TextAlignmentType.Left

				'Set vertical alignment
				styleCell1.VerticalAlignment = TextAlignmentType.Center

				'Set indenting level
				styleCell1.IndentLevel = 1

				'Set top, bottom, left and right borders style
				styleCell1.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
				styleCell1.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Dashed
				styleCell1.Borders(BorderType.TopBorder).LineStyle = CellBorderType.None
				styleCell1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.None

				'Check for last row
				If currentRow = ds.Tables(0).Rows.Count Then
					'Set bottom border style of last row
					styleCell1.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
				End If

				'Set foreground color
				styleCell1.ForegroundColor = color

				'Set background pattern style
				styleCell1.Pattern = BackgroundType.Solid

				'Apply style to current row in first column
				cells(currentRow, 0).SetStyle(styleCell1)

				'Set style for the second column cells
				styleIndex = styles.Add()

				'Set style attribute using style index
				Dim styleCell2 As Style = styles(styleIndex)

				'Copy previous style in new attribute
				styleCell2.Copy(styleCell1)

				'Set left and right border style
				styleCell2.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.None
				styleCell2.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

				'Set horizontal text alignment 
				styleCell2.HorizontalAlignment = TextAlignmentType.Right

				'Set indent level
				styleCell2.IndentLevel = 0

				'Set number format
				styleCell2.Custom = "_(""$""* #,##0.00_);_(""$""* (#,##0.00);_(""$""* ""-""??_);_(@_)"

				'Apply style to current row in second column
				cells(currentRow, 1).SetStyle(styleCell2)

				'Add 1 in current row count
				currentRow += 1
			Next i
		End Sub

		Private Sub GenerateChartSheet(ByVal workbook As Workbook, ByVal ds As DataSet)
			'Generate the second chart sheet
			Dim sheetIndex As Integer = workbook.Worksheets.Add(SheetType.Chart)
			Dim sheet2 As Worksheet = workbook.Worksheets(sheetIndex)

			'Name the sheet
			sheet2.Name = "Pareto Chart"

			'Set chart index
			Dim chartIndex As Integer = sheet2.Charts.Add(ChartType.Column, 0, 0, 0, 0)

			'Get chart type
			Dim chart As Chart = sheet2.Charts(chartIndex)

			'Set chart title text
			chart.Title.Text = "Cost Center"

			'Set chart title font
			chart.Title.TextFont.IsBold = True
			chart.Title.TextFont.Size = 16

			'Set series
			Dim series As String = "Cost Data!B2:B" & (ds.Tables(0).Rows.Count + 1)

			'Series add in chart
			chart.NSeries.Add(series, True)

			'Set series name
			chart.NSeries(0).Name = "Annual Cost"

			'Set category
			chart.NSeries.CategoryData = "Cost Data!A2:A" & (ds.Tables(0).Rows.Count + 1)

			'Legend not shown
			chart.ShowLegend = False

			'Set chart style
			workbook.ChangePalette(Color.FromArgb(255, 255, 200), 53)

			'Set plot area foreground color
			chart.PlotArea.Area.ForegroundColor = Color.FromArgb(255, 255, 200)

			'Set major grid line color
			workbook.ChangePalette(Color.FromArgb(121, 117, 200), 54)
			chart.CategoryAxis.MajorGridLines.Color = Color.FromArgb(121, 117, 200)

			'Set series each point color
			For i As Integer = 0 To chart.NSeries(0).Points.Count - 1
				workbook.ChangePalette(Color.FromArgb(10, 100, 180), 55)
				chart.NSeries(0).Points(i).Area.ForegroundColor = Color.FromArgb(10, 100, 180)
				workbook.ChangePalette(Color.FromArgb(255, 255, 200), 53)
				chart.NSeries(0).Points(i).Border.Color = Color.FromArgb(255, 255, 200)
			Next i
		End Sub


	End Class
End Namespace


