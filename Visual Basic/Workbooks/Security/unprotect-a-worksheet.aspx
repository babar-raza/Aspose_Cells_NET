<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="unprotect-a-worksheet.aspx.vb" Inherits="Workbooks_Security_UnprotectAWorksheet"
	Title="Unprotect a Worksheet - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Unprotect a Worksheet - Aspose.Cells
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
			This online demo explains <b>how to unprotect the worksheet</b> in a workbook using
			<a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			To allow anyone for changing, moving, or deleting data, you can unprotect certain
			worksheet and workbook elements. Aspose.Cells component provides the way to unprotect
			the protected worksheets. The demo makes use of an existing excel file, opens it
			and unprotect the first worksheet of the workbook if the worksheet is protected.
			The sheet is unprotected in such a way that you can change, move and delete anything
			in the first worksheet.
		</p>
		<p>
			You can either open the resultant excel file into <b>MS Excel</b> or save directly
			to your disk to check the results.</p>
		<p>
			Please download the
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/Workbooks/ProtectingWorksheet.xls">ProtectingWorksheet.xls</asp:HyperLink>
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
