<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="True"
	Inherits="CreateListObject" Title="Create List Object - Aspose.Cells Demos" CodeBehind="create-list-object.aspx.vb" %>

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
						List Object - Aspose.Cells</h2>
				</td>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo explains how to <b>Add a List Object</b> to the worksheet using
			<a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			Aspose.Cells allows you to create and maintain lists of various sorts for your need.
			The demo uses an excel spreadsheet file "<b>ListObject.xls</b>" and adds list object
			to the first worksheet in the workbook.
		</p>
		<p>
			Click <b>Process</b>to see how example creates a list object based on template file's
			data. You can either open the output excel file into your <b>MS Excel</b> or save
			directly to your disk.
		</p>
		<p>
			Please download the
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/Workbooks/ListObject.xls">ListObject.xls</asp:HyperLink>
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
