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
Imports System.IO
Imports System.Data.OleDb

Namespace Aspose.Cells.Demos.SmartMarker
	Partial Public Class designer
		Inherits System.Web.UI.Page
		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub btnProcess_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Open the template file through streams
			Dim path As String = MapPath(".")
			path = path.Substring(0, path.LastIndexOf("\")) & "\Designer\SmartMarkerDesigner.xls"
			Dim fs As New FileStream(path, FileMode.Open, FileAccess.Read)
			Dim data(fs.Length - 1) As Byte
			fs.Read(data, 0, data.Length)
			fs.Close()

			'Open/Save the template file through Response object
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader("content-disposition", "attachment;  filename=SmartMarkerDesigner.xls")
			Response.BinaryWrite(data)
			Response.End()
		End Sub

	End Class
End Namespace


