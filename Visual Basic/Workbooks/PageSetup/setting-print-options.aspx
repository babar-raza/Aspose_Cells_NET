<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="setting-print-options.aspx.vb" Inherits="Workbooks_PageSetup_SettingPrintOptions"
	Title="Setting Print Options - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Setting Print Options - Aspose.Cells
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
			This online demo describes how to implement<b> Page Setup Settings</b> using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			Many a times, you would like to configure page setup settings for your worksheets
			to control the printing process. For example, you may need to set <b>Print Area</b>
			of the worksheet, <b>Print Titles</b>, <b>Print Gridlines, Print Row/Column Headings,
				Print Black</b> and <b>Draft Quality, Print Comments, Print Cell Errors and Page Ordering
				</b>etc. All these page setup options are available and supported in Aspose.Cells.
		</p>
		<p>
			The demo makes use of an existing excel file, opens it and implements some page
			setup settings related to printing for the first worksheet of the workbook. It sets
			a specific print area, print specific title rows and columns, print grid lines,
			row / column headings, black and white printing with comments on the worksheet.
			It also sets the printing with draft quality, specifies the order and print cell
			errors.
		</p>
		<p>
			You can either open the resultant excel file into <b>MS Excel </b>or save directly
			to your disk.
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
