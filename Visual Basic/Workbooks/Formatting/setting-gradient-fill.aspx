<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="setting-gradient-fill.aspx.vb" Inherits="Workbooks_Formatting_GradientFill"
	Title="Setting Gradient Fill - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<p>
		<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Setting Gradient Fill - Aspose.Cells</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
				</td>
			</tr>
		</table>
		<div style="text-align: left; font-family: Arial; font-size: small;">
			<p>
				This online demo demonstrates how to set <b>Gradient Fill</b> as a cell's <b>background</b>.
			</p>
			<p>
				Aspose.Cells provides the feature to set the gradient colors as cell's background.
				In this online demo, we will apply gradient fill effect on "<b>A1</b>" cell using
				two colors (<b>Red and Green</b>) to demonstrate the feature. You can either open
				the resultant excel files into <b>MS Excel</b> or save directly to your disk.</p>
			<p>
				<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
					<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
					<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
				</asp:DropDownList>
				<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
			</p>
		</div>
</asp:Content>
