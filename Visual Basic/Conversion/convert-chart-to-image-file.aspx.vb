Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports System.Configuration
Imports System.Collections
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports Aspose.Cells
Imports Aspose.Cells.Drawing
Imports Aspose.Cells.Charts


Partial Public Class Chart2Image
	Inherits System.Web.UI.Page
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Protected Sub CreateStaticReport()
		'Create a new Workbook.
		Dim workbook As New Workbook()
		'Get the first worksheet.
		Dim sheet As Worksheet = workbook.Worksheets(0)
		'Set the name of worksheet
		sheet.Name = "ChartSheet"
		'Get the cells collection in the sheet.
		Dim cells As Cells = workbook.Worksheets(0).Cells
		'Put some values into different cells of the sheet.
		cells("A1").PutValue("Region")
		cells("A2").PutValue("France")
		cells("A3").PutValue("Germany")
		cells("A4").PutValue("England")
		cells("A5").PutValue("Sweden")
		cells("A6").PutValue("Italy")
		cells("A7").PutValue("Spain")
		cells("A8").PutValue("Portugal")
		cells("B1").PutValue("Sale")
		cells("B2").PutValue(70000)
		cells("B3").PutValue(55000)
		cells("B4").PutValue(30000)
		cells("B5").PutValue(40000)
		cells("B6").PutValue(35000)
		cells("B7").PutValue(32000)
		cells("B8").PutValue(10000)

		'Create chart
		Dim chartIndex As Integer = 0
		chartIndex = sheet.Charts.Add(ChartType.Pie, 2, 4, 31, 15)
		Dim chart As Chart = sheet.Charts(chartIndex)

		'Set properties of chart title
		chart.Title.Text = "Sales By Region"
		chart.Title.TextFont.Color = System.Drawing.Color.Blue
		chart.Title.TextFont.IsBold = True
		chart.Title.TextFont.Size = 12

		'Set properties of nseries
		chart.NSeries.Add("B2:B8", True)
		chart.NSeries.CategoryData = "A2:A8"
		chart.NSeries.IsColorVaried = True

		'Set the DataLabels in the chart
		Dim datalabels As DataLabels
		For i As Integer = 0 To chart.NSeries.Count - 1

			datalabels = chart.NSeries(i).DataLabels
			datalabels.Position = LabelPositionType.Center
			datalabels.ShowCategoryName = False
			datalabels.ShowValue = False
			datalabels.ShowPercentage = True
			datalabels.ShowLegendKey = False

		Next i

		'Set the Legend.
		Dim legend As Legend = chart.Legend
		legend.Position = LegendPositionType.Left

		'Create a memory stream object.
		Dim ms As New MemoryStream()
		'Conver the chart to image file.
		chart.ToImage(ms, System.Drawing.Imaging.ImageFormat.Emf)
		'Set Response object to stream the image file.
		Dim data() As Byte = ms.ToArray()
		Me.Context.Response.Clear()
		Me.Context.Response.ContentType = "image/emf"
		Me.Context.Response.AddHeader("content-disposition", "attachment; filename=ChartPic.emf")
		Me.Context.Response.OutputStream.Write(data, 0, data.Length)
		'End response to avoid unneeded html after xls
		Me.Response.End()
	End Sub

End Class
