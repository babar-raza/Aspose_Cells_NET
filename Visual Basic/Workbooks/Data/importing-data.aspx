<%@ Page Language="vb" CodeBehind="importing-data.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.ImportingData" MasterPageFile="~/tpl/Demo.Master"
	Title="Importing Data from different data sources - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tbody>
			<tr>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td class="demos-heading-bg" style="width: 100%;">
					<h2 class="demos-heading-bg">
						Importing Data - Aspose.Cells
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
				Aspose.Cells</a> for .NET to <b>Import Data</b> from a variety of data sources
			to fill the worksheets.
		</p>
		<p>
			Aspose.Cells component is a powerful component, which can be used to create a list
			of data from different data sources including <b>Array</b>, <b>ArrayList</b>, <b>DataColumn</b>,
			<b>DataGrid</b>, <b>DataTable</b>, <b>DataView</b>, <b>Array having Formulas</b>,
			<b>DataReader</b>, <b>Object Array</b> and <b>Two Dimensional Array</b> etc. In
			this demo, you are provided a drop down list and a command button to exercise the
			importing tasks. The drop down list includes the names of all the data sources.
			You may select a data source from it and click the button to generate a workbook
			having the related list of data based on data source.
		</p>
		<p>
			Click <b>Create Report</b> to see how example uses the selected data source to import
			data into Excel worksheet. You can either open the resulting excel file into <b>MS Excel</b>
			or save directly to your disk.
		</p>
		<table class="genericTable" style="font-size: 10pt; font-family: Arial">
			<tr>
				<td>
					Importing Data Source:
				</td>
				<td>
					<asp:DropDownList ID="ImportingDataType" runat="server">
						<asp:ListItem Value="0">Array</asp:ListItem>
						<asp:ListItem Value="1">ArrayList</asp:ListItem>
						<asp:ListItem Value="2">DataColumn</asp:ListItem>
						<%--<asp:ListItem Value="3">DataGrid</asp:ListItem>--%>
						<asp:ListItem Value="4">DataTable</asp:ListItem>
						<asp:ListItem Value="5">DataView</asp:ListItem>
						<asp:ListItem Value="6">FormulaArray</asp:ListItem>
						<asp:ListItem Value="7">FromDataReader</asp:ListItem>
						<asp:ListItem Value="8">ObjectArray</asp:ListItem>
						<asp:ListItem Value="9">TwoDimensionArray</asp:ListItem>
					</asp:DropDownList>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<p>
						<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
							<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
							<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
						</asp:DropDownList>
						<asp:Button ID="btnCreateReport" runat="server" Text="Create Report"></asp:Button>
					</p>
				</td>
			</tr>
		</table>
	</div>
</asp:Content>
