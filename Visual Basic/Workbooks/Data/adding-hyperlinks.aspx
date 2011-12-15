<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="adding-hyperlinks.aspx.vb" Inherits="Workbooks_Data_AddingHyperlinks"
	Title="Adding Hyperlinks to Link Data - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tbody>
			<tr>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td class="demos-heading-bg" style="width: 100%;">
					<h2 class="demos-heading-bg">
						Using Hyperlinks - Aspose.Cells
					</h2>
				</td>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This demo shows how to <b>Add Hyperlinks</b> in your worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			A Hyperlink, <b>a connection between the two areas</b>, is commonly used in the
			internet sites that can be viewed. When the users click a hyperlink, they navigate
			to different location or even an address on the World Wide Web. Aspose.Cells component
			supports hyperlinks. You may add and remove hyperlinks in your worksheet based on
			your need using Aspose.Cells APIs. The demo creates an excel file. The file contains
			two hyperlinks. The first hyperlink placed on <b>A1</b> cell makes you navigate
			to Aspose Site. The second one is an <b>internal hyperlink</b> placed on <b>C1</b>
			cell, when you click on it you will navigate to <b>A10</b> cell of the worksheet.
		</p>
		<p>
			Click <b>Process </b>to see how example creates an excel file that contains two
			hyperlinks. You can either open the resulting excel file into <b>MS Excel</b> or
			save directly to your disk.
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
