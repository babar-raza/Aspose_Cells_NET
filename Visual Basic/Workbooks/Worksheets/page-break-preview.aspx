<%@ Page Language="vb" CodeBehind="page-break-preview.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.PageBreakPreview" MasterPageFile="~/tpl/Demo.Master"
	Title="Normal View and Page Break Preview - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Normal View and Page Break Preview - Aspose.Cells
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
			This online demo exhibits how to display a worksheet in <b>Normal View</b> or <b>Page
				Break Preview</b>&nbsp; using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
					Aspose.Cells</a> for .NET.</p>
		<p>
			All the worksheets can be viewed in two modes: <b>Normal View </b>and <b>Page Break
				Preview</b>. Aspose.Cells component allows you to implement these two modes
			with ease. You can set any type of view according to your need. The demo offers
			you two command buttons <b>Hide</b> and <b>Display</b> to exercise the tasks. When
			you click on Hide button, a workbook is created. It sets the normal view for the
			first worksheet in the workbook. It actually hides the page break preview mode.
			When you click on Display button, the demo creates an excel file and sets the page
			break preview mode for the first worksheet in the workbook. By default worksheet
			mode is normal view.</p>
		<p>
			Click <b>Display</b> to see the page-breaks in a document, click <b>Hide</b> to
			hide the page-breaks. You can either open the resulting excel file into <b>MS Excel</b>
			or save directly to your disk.
		</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="Button1" Text="Display" runat="server"></asp:Button>&nbsp;&nbsp;
			<asp:Button ID="Button2" Text="Hide" runat="server"></asp:Button></p>
	</div>
</asp:Content>
