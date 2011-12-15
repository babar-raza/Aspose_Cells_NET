'///////////////////////////////////////////////////////////////////////
' Copyright (C) 2002-2005 Aspose Pty Ltd.  All rights reserved.

' This file is part of Aspose.Cells. The source code in this file 
' is only intended as a supplement to the documentation, and is provided 
' "as is", without warranty of any kind, either expressed or implied.
'///////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Web
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports Aspose.Cells

Namespace Aspose.Cells.Demos

	Public Class NorthwindPage
		Inherits System.Web.UI.Page

		Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

			'  If you have purchased a License, set license like this: 
			'  License license = new License();
			'	license.SetLicense("Aspose.Cells.lic");
			'  An attempt will be made to find a license file named Aspose.Cells.lic
			'  in the folder that contains the component, in the folder that contains the calling assembly,
			'  in the folder of the entry assembly and then in the embedded resources of the calling assembly.

		End Sub




		#Region "Web Form Designer generated code"
		Overrides Protected Sub OnInit(ByVal e As EventArgs)
			InitializeComponent()
			MyBase.OnInit(e)
		End Sub


		Private Sub InitializeComponent()
'			Me.Load += New System.EventHandler(Me.Page_Load);

		End Sub
		#End Region
	End Class
End Namespace
