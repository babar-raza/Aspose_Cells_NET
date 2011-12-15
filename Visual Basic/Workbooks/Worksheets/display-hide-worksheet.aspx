<%@ Page Language="vb" CodeBehind="display-hide-worksheet.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.DisplayHideWorksheet" MasterPageFile="~/tpl/Demo.Master"
	Title="Hide and Unhide Worksheet - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Hide and Unhide Worksheet - Aspose.Cells
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
			This online demo describes how to <b>Show / Hide</b> worksheets in a workbook using
			<a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			Aspose.Cells component allows you to control the visibility of the worksheets in
			the excel files. The demo offers you two command buttons Hide and Display to exercise
			the tasks. When you click on Hide button, a workbook is created with a default worksheet
			and extra worksheet. It makes the first worksheet invisible in the workbook. When
			you click on Display button, the demo creates an excel file and makes the first
			worksheet visible in the workbook.You can either open the resulting excel file into
			MS Excel or save directly to your disk.
		</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button runat="server" ID="Button1" Text="Display" />&nbsp;
			<asp:Button runat="server" ID="Button2" Text=" Hide " />
		</p>
	</div>
</asp:Content>
