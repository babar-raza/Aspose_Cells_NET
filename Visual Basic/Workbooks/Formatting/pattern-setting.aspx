<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="pattern-setting.aspx.vb" Inherits="Workbooks_Formatting_PatternSetting"
	Title="Pattern Settings - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Pattern Settings - Aspose.Cells</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;">
		<p>
			This online demo exhibits how to implement cell <b>Shading with Patterns</b> using
			<a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			Aspose.Cells component allows you to set cell <b>background color</b> and <b>foreground
				color</b> specified by a <b>pattern</b> to enhance and furnish data or information
			in the cell. The demo creates a simple excel file. It sets different formatting
			involving cell shading with certain colors and pattern setting to <b>B1</b> and
			<b>B2</b> cells in the first worksheet of a workbook. The attern of <b>B2</b> cell
			is <b>DiagonalCrosshatch</b> here.You can either open the resulting excel file into
			<b>MS Excel</b> or save directly to your disk.
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
