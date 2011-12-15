<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="adjusting-rows-and-columns.aspx.vb" Inherits="Workbooks_RowsAndColumns_AdjustingRowsAndColumns"
	Title="Adjusting Row Height and Column Width - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Adjusting Row Height and Column Width - Aspose.Cells
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
			This online demo describes how to set standard row height / standard column width
			of all the rows / columns in a worksheet and customize row height and column width
			of a specific row and column in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			Aspose.Cells component allows you to adjust any row height and column width according
			to your needs. The demo creates an excel file with <b>a standard row height (20)</b>
			and <b>standard column width (20)</b> of all the rows and columns in the first worksheet.
			It then, customizes the width of the <b>first column to 12</b> and the <b>second column
				to 40</b>. Moreover it sets the second row height to 8.
		</p>
		<p>
			You can either open the resultant excel file into <b>MS Excel </b>or save directly
			to your disk to check rows height and columns width of the first worksheet.
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
