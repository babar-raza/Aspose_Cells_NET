<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="data-sorting.aspx.vb" Inherits="Workbooks_Data_DataSorting" Title="Data Sorting - Aspose.Cells Demos" %>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="Server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tbody>
			<tr>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td class="demos-heading-bg" style="width: 100%;">
					<h2 class="demos-heading-bg">
						Data Sorting - Aspose.Cells
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
			This online demo exhibits how to sort the data present in different cells of a worksheet
			using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			<b>Sorting</b> is performed on a list, which is defined as a contiguous group of
			data where the data is displayed in columns. Aspose.Cells allows you to sort Worksheet
			data <b>alphabetically</b> or <b>numerically</b>. It works in the same way as MS
			Excel does to sort data. The demo uses a template file "<b>unsorted.xls</b>" and
			sorts data for data range (<b>A1:B14</b>) in the first worksheet.
		</p>
		<p>
			Click <b>Process </b>to see how example opens an excel file and applies data sorting
			on a range of cells. You can either open the resulting excel file into <b>MS Excel</b>
			or save directly to your disk.
		</p>
		<p>
			Please download
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~\designer\Workbooks\unsorted.xls">unsorted.xls</asp:HyperLink>
			used in this demo.</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
		</p>
	</div>
</asp:Content>
