Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Web
Imports System.Web.SessionState

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for Global.
	''' </summary>
	Public Class [Global]
		Inherits System.Web.HttpApplication
		Public Sub New()
			InitializeComponent()
		End Sub

		Protected Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
#If SITE_BUILD Then
			' NOTE that in production you would want to call 
			' new License().SetLicense("path-to-license-file")
			Try
				Dim lic As New Aspose.Cells.License()
				Aspose.Demos.Common.WebOperationsBridge.InitLicense(lic)

                'Dim lic2 As New Aspose.Cells.GridWeb.License()
                'Aspose.Demos.Common.WebOperationsBridge.InitLicense(lic2)
			Catch
			End Try
#End If
		End Sub

		Protected Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub Application_EndRequest(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		Protected Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)

		End Sub

		#Region "Web Form Designer generated code"
		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
		End Sub
		#End Region
	End Class
End Namespace

