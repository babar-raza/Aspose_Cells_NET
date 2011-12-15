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
Imports System.IO
Imports Aspose.Cells
Imports Aspose.Cells.Drawing
Imports Aspose.Cells.Charts


Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description
	''' </summary>
	Public Class SettingBackgroundImageOfChartSheet
		Inherits System.Web.UI.Page
		Protected Button2 As System.Web.UI.WebControls.Button
		Protected WithEvents Button1 As System.Web.UI.WebControls.Button
		Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

		Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			' Put user code to initialize the page here
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

		Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
			'Create a new workbook
			Dim workbook As New Workbook()

			AddWorksheets(workbook)

			If ddlFileVersion.SelectedItem.Value = "XLS" Then
				'//Save file and send to client browser using selected format
				workbook.Save(HttpContext.Current.Response, "ChartSheet.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
			Else
				workbook.Save(HttpContext.Current.Response, "ChartSheet.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
			End If

			'end response to avoid unneeded html
			HttpContext.Current.Response.End()
		End Sub

		Private Sub AddWorksheets(ByVal workbook As Workbook)
			'//Create a Stream object
			Dim fstream As New FileStream(System.Web.HttpContext.Current.Server.MapPath("~/Image/school.JPG"), FileMode.Open)

			Dim Data(fstream.Length - 1) As Byte

			'//Obtain the file into the array of bytes from streams.
			fstream.Read(Data, 0, Data.Length)

			'Get First Worksheet of the Workbook
			Dim ws As Worksheet = workbook.Worksheets(0)

			'Set Worksheet Type
			ws.Type = SheetType.Chart

			'Set Worksheet background image
			ws.SetBackground(Data)

			'Add new Data Sheet
'INSTANT VB NOTE: The variable data was renamed since Visual Basic will not allow local variables with the same name as parameters or other local variables:
			Dim data_Renamed As Worksheet = workbook.Worksheets.Add("Sheet2")

			'Get data sheet's cells collection
			Dim cells As Cells = data_Renamed.Cells

			'Add Values to cells
			cells("A1").PutValue("Aspose.Cells")

			cells("A2").PutValue("Aspose.Words")

			cells("A3").PutValue("Aspose.PDF")

			cells("B1").PutValue(35)

			cells("B2").PutValue(55)

			cells("B3").PutValue(10)

			'Adding a new chart
			Dim index As Integer = ws.Charts.Add(ChartType.Pie, 5, 0, 15, 5)

			'get newly added chart
			Dim chart As Chart = ws.Charts(index)

			'add nseries of the chart
			chart.NSeries.Add("Sheet2!B1:B3", True)
			chart.NSeries.CategoryData = "Sheet2!A1:A3"

			'show Data Labels
			chart.NSeries(0).DataLabels.ShowCategoryName = True
			chart.NSeries(0).DataLabels.ShowPercentage = True

			'No formatting for Chart Area
			chart.ChartArea.Area.Formatting = FormattingType.None
		End Sub

	End Class
End Namespace


