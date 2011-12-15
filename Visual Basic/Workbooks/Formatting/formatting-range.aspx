<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="formatting-range.aspx.vb" Inherits="Workbooks_Formatting_FormattingRange"
	Title="Formatting Range - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" style="width: 19px">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
			</td>
			<td class="demos-heading-bg" style="width: 100%">
				<h2 class="demos-heading-bg">
					Formatting Range - Aspose.Cells</h2>
			</td>
			<td valign="top" style="width: 19px">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
			</td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;">
		<p>
			This online demo shows how to format a range in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			You can format a range with few lines of code using Aspose.Cells APIs. The demo
			creates an excel file and a range then fill it with blue color. You can either open
			the resultant excel file into <b>MS Excel</b> or save directly to your disk.
		</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
		</p>
	</div>
</asp:Content>
