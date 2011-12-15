<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="calculate-formula.aspx.vb" Inherits="Workbooks_Data_CalculateFormula"
	Title="Calculate Formula - Aspose.Cells Demos" %>

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
						Calculate Formula - Aspose.Cells
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
			This demo explains how to calculate the results of different types of worksheet
			<b>formulas</b> / <b>functions</b> to process data in the spreadsheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			The demo compares the formula / function results of Aspose.Cells with MS Excel.
			Aspose.Cells component supports all the commonly used functions related to different
			categories: <b>Mathematical</b>, <b>String</b>, <b>Statistical</b>, <b>DateTime</b>,
			<b>Logical</b> and <b>Lookup</b> and <b>Reference Functions</b> etc. The demo makes
			use of a template excel file which contains a list of all the formulas / functions
			string of all the categories mentioned. The file also contains some static data
			used in different formulas. The demo retrieves the formulas / functions string and
			calculates the formulas / functions. It also retrieves values from the formulated
			cells and inserts into a column. You can either open the resulting excel file into
			<b>MS Excel</b> or save directly to your disk.
		</p>
		<p>
			Please download
			<asp:HyperLink ID="lnFile" runat="server" NavigateUrl="~\designer\Workbooks\CalculateFormula.xls">CalculateFormula.xls</asp:HyperLink>
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
