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

Partial Public Class EncryptingFile
	Inherits System.Web.UI.Page
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub btnExecute_Click(ByVal sender As Object, ByVal e As EventArgs)
		CreateStaticReport()
	End Sub

	Public Sub CreateStaticReport()
		'Open template.
		Dim path As String = System.Web.HttpContext.Current.Server.MapPath("~")
		path = path.Substring(0, path.LastIndexOf("\"))
		path &= "\designer\book1.xls"

		'Instantiate a new Workbook object.
		Dim workbook As New Workbook(path)

		'Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
		workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128)

		'Use this line if you want to specify XOR Encrytion type.
		'workbook.SetEncryptionOptions(EncryptionType.XOR, 40);

		'Password protect the file.
		workbook.Settings.Password = "007"

		If ddlFileVersion.SelectedItem.Value = "XLS" Then
			'//Save file and send to client browser using selected format
			workbook.Save(HttpContext.Current.Response, "EncryptedBook.xls", ContentDisposition.Attachment, New XlsSaveOptions(SaveFormat.Excel97To2003))
		Else
			workbook.Save(HttpContext.Current.Response, "EncryptedBook.xlsx", ContentDisposition.Attachment, New OoxmlSaveOptions(SaveFormat.Xlsx))
		End If

		'end response to avoid unneeded html
		HttpContext.Current.Response.End()


	End Sub
End Class
