using System;
using System.Data;
using System.Data.OleDb;
using System.Web;

namespace Aspose.Cells.Demos
{
	/// <summary>
	/// Summary description for DBBase.
	/// </summary>
	public class DbBase
	{
		protected System.Data.OleDb.OleDbConnection oleDbConnection1;
		protected System.Data.OleDb.OleDbDataAdapter oleDbDataAdapter1;
		protected System.Data.OleDb.OleDbCommand oleDbSelectCommand1;
		protected System.Data.OleDb.OleDbDataAdapter oleDbDataAdapter2;
		protected System.Data.OleDb.OleDbCommand oleDbSelectCommand2;
		protected DataTable dataTable1;
		protected string path;

		public DbBase(string path)
		{
			this.path = path;
		}

        public string MapPath(string virtualPath)
        {
            return HttpContext.Current.Server.MapPath(virtualPath);
        }

		protected void DBInit()
		{
			this.oleDbConnection1 = new OleDbConnection();
			this.oleDbDataAdapter1 = new OleDbDataAdapter();
			this.oleDbSelectCommand1 = new OleDbCommand();
			this.oleDbDataAdapter2 = new OleDbDataAdapter();
			this.oleDbSelectCommand2 = new OleDbCommand();
			
			this.oleDbConnection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "\\Database\\Northwind.mdb";			

			this.oleDbSelectCommand1.Connection = this.oleDbConnection1;
			this.oleDbDataAdapter1.SelectCommand = this.oleDbSelectCommand1;
			this.oleDbSelectCommand2.Connection = this.oleDbConnection1;
			this.oleDbDataAdapter2.SelectCommand = this.oleDbSelectCommand1;
			
			this.dataTable1 = new DataTable();

		}
	}
}
