/////////////////////////////////////////////////////////////////////////
// Copyright (C) 2002-2005 Aspose Pty Ltd.  All rights reserved.

// This file is part of Aspose.Cells. The source code in this file 
// is only intended as a supplement to the documentation, and is provided 
// "as is", without warranty of any kind, either expressed or implied.
/////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using Aspose.Cells;

namespace Aspose.Cells.Demos
{
	
	public class NorthwindPage : System.Web.UI.Page
	{
		
		private void Page_Load(object sender, System.EventArgs e)
		{

			//  If you have purchased a License, set license like this: 
			//  License license = new License();
			//	license.SetLicense("Aspose.Cells.lic");
			//  An attempt will be made to find a license file named Aspose.Cells.lic
			//  in the folder that contains the component, in the folder that contains the calling assembly,
			//  in the folder of the entry assembly and then in the embedded resources of the calling assembly.
            
		}

		
		
		
		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			InitializeComponent();
			base.OnInit(e);
		}
		
		
		private void InitializeComponent()
		{    
			this.Load += new System.EventHandler(this.Page_Load);
			
		}
		#endregion
	}
}
