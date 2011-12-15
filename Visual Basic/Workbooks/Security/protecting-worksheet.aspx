<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="protecting-worksheet.aspx.vb" Inherits="Workbooks_Security_ProtectingWorksheet"
	Title="Protecting a Worksheet - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Protecting a Worksheet - Aspose.Cells
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
			This online demo explains <b>how to protect the worksheet in a workbook</b> using
			<a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			To prevent anyone from accidentally or deliberately changing, moving, or deleting
			important data, you can protect certain worksheet , workbook elements, with or without
			a password. Aspose.Cells component provides the way to protect your worksheet. You
			may protect <b>Contents</b>, <b>Objects</b> and <b>Scenarios</b> etc. The demo makes
			use of an existing excel file, opens it and protect the first worksheet of the workbook
			without setting a password. The sheet is protected in such a way that <b>you cannot
				change, move and delete anything in the first worksheet</b>. You can either
			open the resultant excel file into <b>MS Excel</b> or save directly to your disk
			to check the protection.
		</p>
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
			<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
		</p>
	</div>
</asp:Content>
