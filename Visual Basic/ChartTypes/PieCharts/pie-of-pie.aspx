<%@ Page Language="vb" Codebehind="pie-of-pie.aspx.vb" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.PieofPie"
	MasterPageFile="~/tpl/Demo.Master" Title="Pie of Pie - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
			<td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
				font-size: large;">
				<h2 class="demos-heading-bg">
					Pie of Pie - Aspose.Cells</h2>
			</td>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo exhibits how to create a Pie of Pie chart in a worksheet using
			<a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			This is a pie chart with user-defined values extracted and combined into a second
			pie. For example, to make small slices easier to see, you can group them together
			as one item in a pie chart and then break down that item in a smaller pie next to
			the main chart. Aspose.Cells is a powerful component, which supports all the standard
			and custom charts to help you display data in more meaningful ways. The demo creates
			a workbook first and inputs the chart source data into the first two columns (A
			and B) of the first worksheet named Data. The first column represents the category
			data (Region) where as the second column represents the sales data representing
			values.
		</p>
		<p>
			The demo creates a pie of pie chart representing Sales By Region into the second
			worksheet named Chart based on the different sale values related to different regions
			in the first worksheet. In the demo, you are provided a sample snapshot of the chart
			and a command button labeled Create Report to create the chart. You can either open
			the resultant excel file into MS Excel or save directly to your disk.
		</p>
		<p>
			Click <b>Create Report</b> to see how demo can set the appearance properties of
			a pie of pie chart.</p>
	</div>
	<table class="genericTable" style="font-family: Arial; font-size: small;">
		<tr>
			<td align="right">
				<img alt="" src="../../Image/PieofPie.jpg" /></td>
			<td valign="top" align="left">
				<table class="genericTable">
					<tr>
						<td>
							Save Format:
						</td>
						<td>
							<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
								<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
								<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td colspan="2">
							<asp:Button ID="btnProcess" runat="server" Text="Create Report"></asp:Button></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</asp:Content>
