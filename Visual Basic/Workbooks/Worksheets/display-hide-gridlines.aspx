<%@ Page Language="vb" CodeBehind="display-hide-gridlines.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.DisplayHideGridlines" MasterPageFile="~/tpl/Demo.Master"
	Title="Display and Hide Grid Lines - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Display and Hide Grid Lines - Aspose.Cells
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
			This online demo explains how to display / hide worksheet gridlines in the workbook
			using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			Many a times, you require to hide gridlines to furnish your worksheets. Aspose.Cells
			component allows you to <b>control the visibility of the worksheet gridlines</b>
			in the excel files. The demo offers you two command buttons <b>Hide</b> and <b>Display</b>
			to exercise the tasks. When you click on Hide button, a workbook is created. It
			makes the gridlines invisible of the first worksheet in the workbook. When you click
			on Display button, the demo creates an excel file and makes the gridlines visible
			of the first worksheet in the workbook. By default worksheet gridlines are visible
			in the workbook. You can either open the resulting excel file into <b>MS Excel</b>
			or save directly to your disk.
		</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button runat="server" ID="Button1" Text="Display" />
			<asp:Button runat="server" ID="Button2" Text=" Hide " />
		</p>
	</div>
</asp:Content>
