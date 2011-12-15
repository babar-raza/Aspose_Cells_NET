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
	''' Summary description for SchoolReport.
	''' </summary>
	Public Class SchoolReport
		Inherits System.Web.UI.Page
		Protected WithEvents Button1 As System.Web.UI.WebControls.Button
		Protected ListBox1 As System.Web.UI.WebControls.ListBox
		Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

		Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			' Put user code to initialize the page here

			If (Not IsPostBack) Then
				CreateList()
			Else
				If Me.Request.Params.Count > 0 Then
					Dim param As String = Me.Request.Params(0)
					'Instantiate a workbook
					Dim workbook As New Workbook()
					CreateStaticReport(workbook)
					CreateDynamicReport(workbook)


					If ddlFileVersion.SelectedItem.Value = "XLS" Then
						'//Save file and send to client browser using selected format
						workbook.Save(HttpContext.Current.Response, "ReportCard.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
					Else
						workbook.Save(HttpContext.Current.Response, "ReportCard.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
					End If

					'end response to avoid unneeded html
					HttpContext.Current.Response.End()
				End If

			End If

		End Sub

		#Region "Web Form Designer generated code"
		Overrides Protected Sub OnInit(ByVal e As EventArgs)
			'
			' CODEGEN: This call is required by the ASP.NET Web Form Designer.
			'
			InitializeComponent()
			MyBase.OnInit(e)
		End Sub

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
'			Me.Button1.Click += New System.EventHandler(Me.Button1_Click);
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region

		''' <summary>
		''' Creates student list.
		''' </summary>
		Private Sub CreateList()
			'Creates student list from data in an Workbook file. 
			'In a real world application, all kind of data sources can be used.

			'Open the template
			Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
			path = path.Substring(0, path.LastIndexOf("\"))
			path &= "\designer\SchoolData.xls"


			'string path = MapPath("~/designer/SchoolData.xls");

			Dim dataFile As String = path
			Dim workbook As New Workbook(dataFile)


			'Get the cells collection in the first worksheet
			Dim cells As Cells = workbook.Worksheets(0).Cells
			'Export the sheet data to a multi-dimensional array
			Dim nameList(,) As Object = cells.ExportArray(1, 0, cells.MaxDataRow, 2)
			'Fill the list box
			For i As Integer = 0 To nameList.Length / 2 - 1
				Me.ListBox1.Items.Add(nameList(i, 0).ToString() & " " & nameList(i, 1).ToString())
			Next i
			'Set first element as a selected item
			Me.ListBox1.SelectedIndex = 0
		End Sub

		Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
			'Redirect to the same page with some parameter
			Response.Redirect("SchoolReport.aspx?Data=abc")
		End Sub

		Private Sub CreateStaticReport(ByVal workbook As Workbook)
			'Sets default font
			Dim style As Style = workbook.DefaultStyle
			style.Font.Name = "Tahoma"
			workbook.DefaultStyle = style

			'Get the first worksheet in the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Name the worksheet
			sheet.Name = "Report Card"
			'Make the gridlines insible for the sheet
			sheet.IsGridlinesVisible = False

			AddImageAndChart(workbook)

			'Add a new worksheet to the workbook
			Dim index As Integer = workbook.Worksheets.Add()
			'Get the sheet
			sheet = workbook.Worksheets(index)
			'Name the sheet
			sheet.Name = "Grade Table"
			'Make the gridlines invisible for the worksheet
			sheet.IsGridlinesVisible = False


			SetRowColumn(workbook)
			CreateOutline(workbook)
			CreateCellsFormatting(workbook)
			CreateStaticData(workbook)



		End Sub

		Private Sub SetRowColumn(ByVal workbook As Workbook)
			'Get the cells in the first worksheet
			Dim cells As Cells = workbook.Worksheets(0).Cells
			'Set the height of the first 22 rows
			For i As Integer = 0 To 21
				cells.SetRowHeight(i, 13.5)
			Next i

			'Set the height of the next row
			cells.SetRowHeight(22, 6)
			'Set the height for the 24th row
			cells.SetRowHeight(23, 13.5)
			'Set the row height for the next 5 rows
			For i As Integer = 24 To 28
				cells.SetRowHeight(i, 22.5)
			Next i
			'Set the row height for the next two rows
			cells.SetRowHeight(29, 13.5)
			cells.SetRowHeight(30, 6)

			'Set the row height for the 32-34 rows
			For i As Integer = 31 To 33
				cells.SetRowHeight(i, 13.5)
			Next i


			'Set the columns widths for first four (A-D) columns
			cells.SetColumnWidth(0, 1.86)
			cells.SetColumnWidth(1, 1.86)
			cells.SetColumnWidth(2, 19)
			cells.SetColumnWidth(3, 15.14)

			'Set the column widths for 5-10 columns
			For column As Byte = 4 To 9
				cells.SetColumnWidth(column, 5)
			Next column

			'Set the column widths for 12-13 columns
			cells.SetColumnWidth(11, 15.43)
			cells.SetColumnWidth(12, 2)

			'Get the third worksheet cells
			cells = workbook.Worksheets(2).Cells
			'Set the row height for the second row
			cells.SetRowHeight(1, 15.75)
			'Set the column widths for first two columns
			cells.SetColumnWidth(0, 2)
			cells.SetColumnWidth(1, 11.86)
			'Set the Column widths for 3-15 columns
			For i As Integer = 2 To 14
				cells.SetColumnWidth(i, 5.14)
			Next i
		End Sub

		Private Sub CreateOutline(ByVal workbook As Workbook)
			'Get the first worksheet cells
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Create a range and its outline borders
			Dim range As Range = cells.CreateRange("B2", "M22")
			range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))
			range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))
			range.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))
			range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))

			'Create a range and its outline borders
			range = cells.CreateRange("B24", "M30")
			range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))
			range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))
			range.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))
			range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))

			'Create a range and its outline borders
			range = cells.CreateRange("B32", "M34")
			range.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))
			range.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))
			range.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))
			range.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Medium, Color.FromArgb(0, 0, 128))
		End Sub

		Private Sub CreateCellsFormatting(ByVal workbook As Workbook)
			'Creates cell formatting on the first worksheet

			'Create a style object
			Dim style As Style = workbook.Styles(workbook.Styles.Add())
			'Sets font attributes
			style.Font.IsBold = True
			style.Font.Size = 10

			'Set the style to some cells
			Dim cells As Cells = workbook.Worksheets(0).Cells
			cells("K4").SetStyle(style)
			cells("E11").SetStyle(style)
			cells("E12").SetStyle(style)

			'Create the style object
			style = workbook.Styles(workbook.Styles.Add())
			'Sets font attributes
			style.Font.IsBold = True
			style.HorizontalAlignment = TextAlignmentType.Center

			'Sets borders
			style.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'Apply style to some cells (C15-L15)
			Dim startColumn As Integer = CellsHelper.ColumnNameToIndex("C")
			Dim endColumn As Integer = CellsHelper.ColumnNameToIndex("L")
			For i As Integer = startColumn To endColumn
				cells(14, i).SetStyle(style)
			Next i

			'Create the style object and set borders
			style = workbook.Styles(workbook.Styles.Add())
			style.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'Sets foreground color
			style.ForegroundColor = Color.FromArgb(&HFF, &HFF, &HCC)
			style.Pattern = BackgroundType.Solid

			'Apply style to some specific cells in rows
			For i As Integer = startColumn To endColumn
				cells(15, i).SetStyle(style)
				cells(17, i).SetStyle(style)
				cells(19, i).SetStyle(style)
			Next i

			'Create the style and set borders
			style = workbook.Styles(workbook.Styles.Add())
			style.Borders(BorderType.TopBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.LeftBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.RightBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin

			'Sets foreground color
			style.ForegroundColor = Color.FromArgb(&HCC, &HFF, &HCC)
			style.Pattern = BackgroundType.Solid

			'Apply style to some cells in rows
			For i As Integer = startColumn To endColumn
				cells(16, i).SetStyle(style)
				cells(18, i).SetStyle(style)
				cells(20, i).SetStyle(style)
			Next i

			'Create the style object and set borders
			style = workbook.Styles(workbook.Styles.Add())
			style.Borders(BorderType.BottomBorder).Color = Color.FromArgb(0, 0, 128)
			style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin

			'Apply style to some cells in rows
			For i As Integer = startColumn To endColumn
				cells(24, i).SetStyle(style)
				cells(25, i).SetStyle(style)
				cells(26, i).SetStyle(style)
				cells(27, i).SetStyle(style)
				cells(28, i).SetStyle(style)
			Next i

			'Apply style to some cells in a row
			For i As Integer = 4 To 9
				cells(32, i).SetStyle(style)
			Next i
			cells(32, 11).SetStyle(style)

			'Apply some custom number style in a column's cells
			For i As Integer = 15 To 20
				Dim style1 As Aspose.Cells.Style = New Style()
				style.Custom = "0"
				cells(i, 10).SetStyle(style1)
			Next i

			'Get the second worksheet cells collection
			cells = workbook.Worksheets(2).Cells

			'Create the style object
			style = workbook.Styles(workbook.Styles.Add())
			'Specify the forground color and font attributes
			style.ForegroundColor = Color.FromArgb(128, 0, 0)
			style.Pattern = BackgroundType.Solid
			style.Font.Color = Color.FromArgb(255, 255, 153)
			style.Font.Size = 12
			style.Font.IsBold = True

			'Apply the style to some cells in the second row
			For i As Integer = 1 To 14
				cells(1, i).SetStyle(style)
			Next i

			'Create the style object and set borders
			style = workbook.Styles(workbook.Styles.Add())
			style.Borders(BorderType.TopBorder).Color = Color.Black
			style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.BottomBorder).Color = Color.Black
			style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.LeftBorder).Color = Color.Black
			style.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.RightBorder).Color = Color.Black
			style.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
			'Specify the alignment type
			style.HorizontalAlignment = TextAlignmentType.Center
			'Set the font attribute
			style.Font.IsBold = True
			'Apply style to B3 and B4 cells
			cells("B3").SetStyle(style)
			cells("B4").SetStyle(style)

			'Create the style and set borders
			style = workbook.Styles(workbook.Styles.Add())
			style.Borders(BorderType.TopBorder).Color = Color.Black
			style.Borders(BorderType.TopBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.BottomBorder).Color = Color.Black
			style.Borders(BorderType.BottomBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.LeftBorder).Color = Color.Black
			style.Borders(BorderType.LeftBorder).LineStyle = CellBorderType.Thin
			style.Borders(BorderType.RightBorder).Color = Color.Black
			style.Borders(BorderType.RightBorder).LineStyle = CellBorderType.Thin
			'Set alignment type
			style.HorizontalAlignment = TextAlignmentType.Center
			'Apply style to some range of cells
			For i As Integer = 2 To 3
				For j As Integer = 2 To 14
					cells(i, j).SetStyle(style)
				Next j
			Next i
		End Sub

		Private Sub CreateStaticData(ByVal workbook As Workbook)
			'Get the cells in the first worksheet
			Dim cells As Cells = workbook.Worksheets(0).Cells

			'Get Style Object 
			Dim style As Aspose.Cells.Style = cells("E4").GetStyle()

			'Put values, apply formula(s) to the cells with style formatting
			cells("E4").PutValue("ASPOSE School District")
			style.Font.IsBold = True
			style.Font.Size = 12
			cells("E4").SetStyle(style)
			cells("K4").PutValue("Progress Report")
			cells("E5").PutValue("Suite 180, 9 Crofts Avenue")
			cells("K5").PutValue("Date:")
			style.Font.IsBold = True
			cells("K5").SetStyle(style)
			cells("L5").Formula = "=Now()"
			style.Custom = "[$-409]mmmm d, yyyy;@"
			cells("L5").SetStyle(style)
			cells("E6").PutValue("Hurstville, NSW, 2220")
			cells("E8").PutValue("Phone: 888.277.6734")
			cells("E9").PutValue("Fax: 866.810.9465")
			cells("E11").PutValue("Student Name")
			cells("E12").PutValue("Student SSN")
			cells("C15").PutValue("Class Name")
			cells("D15").PutValue("Teacher")
			cells("E15").PutValue("1st")
			cells("F15").PutValue("2nd")
			cells("G15").PutValue("3rd")
			cells("H15").PutValue("4th")
			cells("I15").PutValue("5th")
			cells("J15").PutValue("6th")
			cells("K15").PutValue("Final")
			cells("L15").PutValue("Letter Grade")
			cells("C16").PutValue("English")
			cells("C17").PutValue("Math")
			cells("C18").PutValue("Social Studies")
			cells("C19").PutValue("Science")
			cells("C20").PutValue("Art")
			cells("C21").PutValue("Physical Education")
			cells("C24").PutValue("Note")
			style.Font.IsBold = True
			cells("C24").SetStyle(style)
			cells("D33").PutValue("Parent Signature:")
			cells("D33").SetStyle(style)
			cells("K33").PutValue("Date")
			cells("K33").SetStyle(style)
			cells = workbook.Worksheets(2).Cells
			cells("B2").PutValue("Grade Table")
			cells("B3").PutValue("Average")
			cells("C3").PutValue(0)
			cells("D3").PutValue(60)
			cells("E3").PutValue(63)
			cells("F3").PutValue(67)
			cells("G3").PutValue(70)
			cells("H3").PutValue(73)
			cells("I3").PutValue(77)
			cells("J3").PutValue(80)
			cells("K3").PutValue(83)
			cells("L3").PutValue(87)
			cells("M3").PutValue(90)
			cells("N3").PutValue(93)
			cells("O3").PutValue(97)
			cells("B4").PutValue("Letter Grade")
			cells("C4").PutValue("F")
			cells("D4").PutValue("D-")
			cells("E4").PutValue("D")
			cells("F4").PutValue("D+")
			cells("G4").PutValue("C-")
			cells("H4").PutValue("C")
			cells("I4").PutValue("C+")
			cells("J4").PutValue("B-")
			cells("K4").PutValue("B")
			cells("L4").PutValue("B+")
			cells("M4").PutValue("A-")
			cells("N4").PutValue("A")
			cells("O4").PutValue("A+")

		End Sub

		Private Sub AddImageAndChart(ByVal workbook As Workbook)
			'Get the image file path
			Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
			path = path.Substring(0, path.LastIndexOf("\"))
			path &= "\Image\School.jpg"


			Dim imageFile As String = path
			'Add image to the first worksheet
			Dim index As Integer = workbook.Worksheets(0).Pictures.Add(1, 1, imageFile)
			Dim pic As Picture = workbook.Worksheets(0).Pictures(index)
			pic.Left = 2
			pic.Top = 2

			'Add a chart worksheet type
			index = workbook.Worksheets.Add(SheetType.Chart)
			'Get the worksheet
			Dim sheet As Worksheet = workbook.Worksheets(index)
			'Set the name
			sheet.Name = "Grade Chart"
			'Set the scalling factor
			sheet.Zoom = 90


			'Add a new bar chart to the worksheet            
			Dim chart As Chart = sheet.Charts(sheet.Charts.Add(ChartType.Bar, 0, 0, 0, 0))

			'Set the nseries data range
			chart.NSeries.Add("'Report Card'!E16:H21", False)
			'Name the series
			For i As Integer = 0 To chart.NSeries.Count - 1
				chart.NSeries(i).Name = "='Report Card'!C" & (16 + i).ToString()
			Next i

			'Set the legend position to bottom on the chart
			chart.Legend.Position = LegendPositionType.Bottom
		End Sub

		Private Sub CreateDynamicReport(ByVal workbook As Workbook)
			'Get the template file path
			Dim path As String = MapPath("~")
			path = path.Substring(0, path.LastIndexOf("\"))
			Dim dataFile As String = path & "\Designer\SchoolData.xls"
			'Get the selected list box item
			Dim name As String = Me.ListBox1.SelectedItem.Text
			'Split the array
			Dim nameArray() As String = name.Split(" "c)
			'Get the first worksheet cells
			Dim cells As Cells = workbook.Worksheets(0).Cells
			'Put the selected value (in the list box) to the cell
			cells("H11").PutValue(name)

			'Instantiate a workbook
			Dim dataWorkbook As New Workbook(dataFile)

			'Get teachers' name
			Dim teachers(dataWorkbook.Worksheets.Count - 1) As String
			For i As Integer = 0 To teachers.Length - 1
				teachers(i) = dataWorkbook.Worksheets(i).Cells("L1").StringValue
			Next i

			'Put teachers' name into output workbook
			cells.ImportArray(teachers, 15, 3, True)


			'Get / Set students data
			Dim cell As Cell = Nothing
			Dim dataSheet As Worksheet = dataWorkbook.Worksheets(dataWorkbook.Worksheets.Count - 1)
			Do
				cell = dataSheet.Cells.FindString(nameArray(0), cell)
				If cell IsNot Nothing Then
					If dataSheet.Cells(cell.Row, cell.Column + 1).StringValue = nameArray(1) Then
						cells("H12").PutValue(dataSheet.Cells(cell.Row, cell.Column + 2).Value)
						Exit Do
					End If
				Else
					Exit Do
				End If
			Loop

			For i As Integer = 0 To dataWorkbook.Worksheets.Count - 2
				Dim studentData As DataTable = dataWorkbook.Worksheets(i).Cells.ExportDataTable(1, 0, cells.MaxDataRow, 8)
				For Each row As DataRow In studentData.Rows
					If row(0).ToString() = nameArray(0) AndAlso row(1).ToString() = nameArray(1) Then
						For j As Integer = 2 To row.ItemArray.Length - 1
							cells(15 + i, j + 2).PutValue(row(j))
						Next j
					End If
				Next row
			Next i

			'Specify some formulas for Marks Average and Grade
			For i As Integer = 15 To 20
				cells(i, 10).Formula = "=AVERAGE(E" & (i + 1).ToString() & ":J" & (i + 1).ToString() & ")"
				cells(i, 11).Formula = "=IF(K" & (i + 1).ToString() & "<>"""",HLOOKUP(K" & (i + 1).ToString() & ",'Grade Table'!$C$3:$O$4,2),"""")"
			Next i

			'Calculate all the formulas
			workbook.CalculateFormula()

			Dim courseIndex As Integer = -1
			Dim minScore As Double = -1
			For i As Integer = 15 To 20
				Dim score As Double = cells(i, 10).DoubleValue
				If score < 80 Then
					If minScore < 0 Then
						minScore = score
						courseIndex = i
					ElseIf score < minScore Then
						minScore = score
						courseIndex = i
					End If
				End If
			Next i

			'Specify some notes for the marksheet
			If courseIndex <> -1 Then
				Dim course As String = cells(courseIndex, 2).StringValue
				Dim note As String = "{0} seems to be having difficulties with {1} projects.  We offer after school tutoring sessions"

				note = String.Format(note, nameArray(0), course)
				cells("C25").PutValue(note)
				cells("C26").PutValue("which may be helpful.")
			End If

		End Sub
	End Class
End Namespace


