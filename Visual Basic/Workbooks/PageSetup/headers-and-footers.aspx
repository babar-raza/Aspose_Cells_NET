<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="headers-and-footers.aspx.vb" Inherits="Workbooks_PageSetup_HeadersAndFooters"
	Title="Setting Headers and Footers - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Setting Headers and Footers - Aspose.Cells
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
			This online demo explains <b>how to implement page setup settings related to headers
				and footers</b> using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
					Aspose.Cells</a> for .NET. </font>
		</p>
		<p>
			Sometimes, you would like to configure page setup settings for your worksheets to
			control the printing process. For example, you may need to set <b>Page Headers and Footers</b>
			in your worksheet. These page setup settings are supported by Aspose.Cells component.
			The component allows you to add headers and footers to the worksheets at runtime.
			To add headers and footers at runtime, the component provides special <b>APIs</b>
			and some <b>Script Commands</b> to control the formatting of headers and footers.
			And you need to understand the vocabulary of the script commands. For reference,
			Please see <a href="http://www.aspose.com/documentation/file-format-components/aspose.cells-for-.net-and-java/setting-headers-and-footers.html">
				Headers &amp; Footers</a>.
		</p>
		<p>
			The demo makes use of an existing excel file, opens it and implements date as a
			header and time as a footer in the first worksheet of the workbook. You can either
			open the output excel file into <b>MS Excel</b> or save directly to your disk and
			use <b>Print Preview</b> option to check the results in <b>MS Excel</b>.</p>
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
