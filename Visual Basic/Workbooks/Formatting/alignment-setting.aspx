<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="alignment-setting.aspx.vb" Inherits="Workbooks_Formatting_AlignmentSetting"
	Title="Alignment Settings - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Alignment Settings - Aspose.Cells</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;">
		<p>
			This online demo demonstrates some alignment tasks related to text: e.g., <b>Text Alignment</b>,
			<b>Orientation</b>, <b>Text Control</b> and <b>Indentation</b>.
		</p>
		<p>
			Aspose.Cells gives you the flexibility to perform all the <b>Alignment Tasks</b>.
			You can align the text <b>horizontally</b> and <b>vertically</b>, you may rotate
			the text to a certain degree, you are allowed to <b>shrink the text to fit</b> in
			a cell, you can set the indentation level of the text in a cell, you may <b>wrap the
				text</b> and set the <b>text direction</b> etc. The demo makes use of an excel
			file and performs all the important alignment tasks mentioned above.
		</p>
		<p>
			You can either open the resultant excel files into <b>MS Excel</b> or save directly
			to your disk.</p>
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
