<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="conditional-formatting.aspx.vb" Inherits="Workbooks_Formatting_ConditionalFormatting"
	Title="Conditional Formatting - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" style="width: 19px">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
			</td>
			<td class="demos-heading-bg" style="width: 100%">
				<h2 class="demos-heading-bg">
					Conditional Formatting - Aspose.Cells</h2>
			</td>
			<td valign="top" style="width: 19px">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
			</td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;">
		<p>
			This online demo exhibits how to apply <b>Conditional Formatting</b> on different
			cells using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			Conditional Formatting is an advance feature in <b>Microsoft Excel</b> that allows
			you to apply formats to a <b>Range of cells</b>, and have that formatting change
			depending on the value of the cell or the <b>Value of a Formula</b>. For example,
			you can have a <b>Cell Background</b> only when the value of the cell is <b>greater
				than 50</b>. When the value of the cell meets the format condition, the format
			you select is applied to the cell. If the value of the cell does not meet the format
			condition, the cell's default formatting is used. You can either open the resultant
			excel files into <b>MS Excel</b> or save directly to your disk to check the results.</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" /></p>
	</div>
</asp:Content>
