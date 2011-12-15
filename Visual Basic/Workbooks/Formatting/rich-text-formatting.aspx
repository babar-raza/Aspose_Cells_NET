<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="rich-text-formatting.aspx.vb" Inherits="Workbooks_Formatting_RichTextFormatting"
	Title="Rich Text Formatting - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" style="width: 19px">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
			</td>
			<td class="demos-heading-bg" style="width: 100%">
				<h2 class="demos-heading-bg">
					Rich Text Formatting - Aspose.Cells</h2>
			</td>
			<td valign="top" style="width: 19px">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
			</td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;">
		<p>
			This online demo exhibits how to apply <b>Rich Text Formatting</b> using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			You can use rich text formatting to display data in more enhanced and elaborated
			way. You can use different options like <b>bold</b>,<b>italic</b>,<b>color</b> and
			different <b>style</b> options to make important data more distinguish. In this
			demo we will apply rich text formatting on a cell's data to show how easy it is
			to achieve the rich text formatting using Aspose.Cells. You can either open the
			resultant excel files into <b>MS Excel</b> or save directly to your disk to check
			the results.</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
		</p>
	</div>
</asp:Content>
