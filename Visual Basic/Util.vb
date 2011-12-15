Imports Microsoft.VisualBasic
Imports System

Namespace Aspose.Cells.GridWeb.DemosCS
	''' <summary>
	''' some common Methods for demos
	''' </summary>
	Public Class Util
		Public Shared Sub ShowMessage(ByVal page As System.Web.UI.Page, ByVal msg As String)
			Dim script As String
			script = "<script language='javascript'>alert('" & msg & "')</script>"
			page.ClientScript.RegisterClientScriptBlock(page.GetType(), "Util.alertMessage", script)
			'page.RegisterClientScriptBlock("Util.alertMessage", script);      
		End Sub
	End Class
End Namespace
