<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="advanced-protection.aspx.vb" Inherits="Workbooks_Security_AdvancedProtection"
	Title="Advanced Protection Settings Since MS-Excel XP - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Advanced Protection Settings Since MS-Excel XP - Aspose.Cells
					</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo shows how to customizes the <b>Advanced Protection Settings</b>
			of the worksheet in a workbook using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			These settings can only be implemented in <b>MS Excel XP or greater versions</b>.
			Aspose.Cells component provides some useful APIs to protect the worksheet even at
			the root level. There are a number of protection options available which you may
			set according to your need. For example, you may implement: whether the user is
			allowed to filter data, whether the user is allowed to do formatting of cells, rows
			/ columns, whether the user can add hyperlinks to the cells, whether the user can
			do sorting of data in the worksheet, whether the user is authorized to insert new
			rows / columns or select the locked cells etc.
		</p>
		<p>
			The demo makes use of an <b>existing excel file</b>, opens it and <b>implements some
				advanced protection settings in the first worksheet</b> of the file. When you
			will try to edit any cell, it will show you a message that cell is protected.
		</p>
		<p>
			You can either open the resultant excel file into <b>MS Excel</b> or save directly
			to your disk to check the results.</p>
		<p>
			Please download the
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/book1.xls">book1.xls</asp:HyperLink>
			used in this demo.
		</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" /></p>
	</div>
</asp:Content>
