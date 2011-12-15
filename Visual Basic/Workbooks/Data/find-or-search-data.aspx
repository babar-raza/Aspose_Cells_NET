<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="find-or-search-data.aspx.vb" Inherits="Workbooks_Data_FindOrSearchData"
	Title="Find or Search Data - Aspose.Cells Demos" %>

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
						Find or Search Data - Aspose.Cells
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
			This online demo demonstrates the ability of <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET to find or <b>Search Cells</b>, containing specified
			data in worksheets.
		</p>
		<p>
			Aspose.Cells provides an easy way to <b>find individual data</b> or <b>a list of data</b>.
			It will even allow you to look for <b>a partial data</b> if you are not sure about
			how the data is completely spelled. You may find <b>string data</b>, <b>numeric data</b>
			and <b>search formulas</b>. The example creates an excel file from the scratch and
			inputs some values into different cells in a worksheet, apply some formulas and
			find data in the cells.You can either open the resulting excel file into <b>MS Excel</b>
			or save directly to your disk.
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
