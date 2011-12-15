Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Web

Namespace Aspose.Cells.Demos
	''' <summary>
	''' Summary description for DBBase.
	''' </summary>
	Public Class DbBase
		Protected oleDbConnection1 As System.Data.OleDb.OleDbConnection
		Protected oleDbDataAdapter1 As System.Data.OleDb.OleDbDataAdapter
		Protected oleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
		Protected oleDbDataAdapter2 As System.Data.OleDb.OleDbDataAdapter
		Protected oleDbSelectCommand2 As System.Data.OleDb.OleDbCommand
		Protected dataTable1 As DataTable
		Protected path As String

		Public Sub New(ByVal path As String)
			Me.path = path
		End Sub

		Public Function MapPath(ByVal virtualPath As String) As String
			Return HttpContext.Current.Server.MapPath(virtualPath)
		End Function

		Protected Sub DBInit()
			Me.oleDbConnection1 = New OleDbConnection()
			Me.oleDbDataAdapter1 = New OleDbDataAdapter()
			Me.oleDbSelectCommand1 = New OleDbCommand()
			Me.oleDbDataAdapter2 = New OleDbDataAdapter()
			Me.oleDbSelectCommand2 = New OleDbCommand()

			Me.oleDbConnection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & "\Database\Northwind.mdb"

			Me.oleDbSelectCommand1.Connection = Me.oleDbConnection1
			Me.oleDbDataAdapter1.SelectCommand = Me.oleDbSelectCommand1
			Me.oleDbSelectCommand2.Connection = Me.oleDbConnection1
			Me.oleDbDataAdapter2.SelectCommand = Me.oleDbSelectCommand1

			Me.dataTable1 = New DataTable()

		End Sub
	End Class
End Namespace
