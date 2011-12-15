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
Imports Aspose.Cells

Partial Public Class Workbooks_PageSetup_SettingPrintOptions
	Inherits System.Web.UI.Page
	Protected ddlFileVersion As System.Web.UI.WebControls.DropDownList

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Open template
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
		path = path.Substring(0, path.LastIndexOf("\"))
		path &= "\designer\book1.xls"


		Dim workbook As New Workbook(path)


		Dim pageSetup As PageSetup = workbook.Worksheets(0).PageSetup
		'Specify the cells range (from A1 cell to B2 cell) of the print area
		pageSetup.PrintArea = "A1:G5"

		'Define column numbers A & B as title columns
		pageSetup.PrintTitleColumns = "$A:$B"

		'Define row numbers 1 & 2 as title rows
		pageSetup.PrintTitleRows = "$1:$2"

		'Allow to print gridlines
		pageSetup.PrintGridlines = True

		'Allow to print row/column headings
		pageSetup.PrintHeadings = True

		'Allow to print worksheet in black & white mode
		pageSetup.BlackAndWhite = True

		'Allow to print comments as displayed on worksheet
		pageSetup.PrintComments = PrintCommentsType.PrintInPlace

		'Allow to print worksheet with draft quality
		pageSetup.PrintDraft = True

		'Allow to print cell errors 
		pageSetup.PrintErrors = PrintErrorsType.PrintErrorsBlank

		'Set the printing order of the pages to over then down
		pageSetup.Order = PrintOrderType.DownThenOver

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "SettingPrintOptions.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "SettingPrintOptions.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub
End Class
