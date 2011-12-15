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

Partial Public Class Workbooks_Security_AdvancedProtection
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


		'Instantiate a workbook
		Dim workbook As New Workbook(path)



		'Get the first worksheet in the workbook
		Dim worksheet As Worksheet = workbook.Worksheets(0)
		'Get the protection in the sheet
		Dim protection As Protection = worksheet.Protection

		'Restricting users to delete columns of the worksheet
		protection.AllowDeletingColumn = False

		'Restricting users to delete row of the worksheet
		protection.AllowDeletingRow = False

		'Restricting users to edit contents of the worksheet
		protection.AllowEditingContent = False

		'Allowing users to edit objects of the worksheet
		protection.AllowEditingObject = True

		'Allowing users to edit scenarios of the worksheet
		protection.AllowEditingScenario = True

		'Restricting users to filter
		protection.AllowFiltering = False

		'Allowing users to format cells of the worksheet
		protection.AllowFormattingCell = True

		'Allowing users to format rows of the worksheet
		protection.AllowFormattingRow = True

		'Allowing users to insert columns in the worksheet
		protection.AllowInsertingColumn = True

		'Allowing users to insert hyperlinks in the worksheet
		protection.AllowInsertingHyperlink = True

		'Allowing users to insert rows in the worksheet
		protection.AllowInsertingRow = True

		'Allowing users to select locked cells of the worksheet
		protection.AllowSelectingLockedCell = True

		'Allowing users to select unlocked cells of the worksheet
		protection.AllowSelectingUnlockedCell = True

		'Allowing users to sort
		protection.AllowSorting = True

		'Allowing users to use pivot tables in the worksheet
		protection.AllowUsingPivotTable = True

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "AdvancedProtection.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "AdvancedProtection.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()
	End Sub

End Class
