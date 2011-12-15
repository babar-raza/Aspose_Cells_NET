<%@ Page Language="vb" Codebehind="pyramid-percent-stacked.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.PyramidPercentStacked" MasterPageFile="~/tpl/Demo.Master"
	Title="Pyramid Percent Stacked - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
			<td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
				font-size: large;">
				<h2 class="demos-heading-bg">
					Pyramid Percent Stacked - Aspose.Cells</h2>
			</td>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo exhibits how to create a Pyramid 100% Stacked Bar/Column chart
			in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			This type of chart compares the percentage each value contributes to a total across
			categories. The bars/columns in these types of chart are represented by pyramid
			shapes. Aspose.Cells is a powerful component, which supports all the standard and
			custom charts to help you display data in more meaningful ways. The demo creates
			a workbook first and inputs some chart related data into the first five columns
			(A, B, C, D and E) of the first worksheet named Data. The first column represents
			the product names (Product1, Product2 and Product3) where as the second, third,
			fourth and fifth columns represent percentage values involving different quarters
			(Qurarter1, Quarter2, Quarter3 and Quarter4).
		</p>
		<p>
			The demo creates a pyramid 100% stacked chart titled Product contribution to total
			sales into the second worksheet named Chart based on the different product values
			related to different quarters in the first worksheet. In the demo, you are provided
			a sample snapshot of the chart, a drop down list that represents the chart type
			(Pyramid100PercentStacked and PyramidBar100PercentStacked) and a command button
			labeled Create Report to create and exercise the chart based on your selection from
			the drop down list. You can either open the resultant excel file into <b>MS Excel</b>
			or save directly to your disk to check the results.
		</p>
		<p>
			Click <b>Create Report</b> to see how demo can set the appearance properties of
			a pyramid percent stacked or percent stacked bar chart.</p>
	</div>
	<table class="genericTable" style="font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" align="right">
				<img  alt="" src="../../Image/PyramidPercentStacked.jpg"/></td>
			<td valign="top" align="left">
				<table class="genericTable">
					<tr>
						<td>
							Chart Type:
						</td>
						<td>
							<asp:DropDownList runat="server" ID="ChartTypeList">
								<asp:ListItem Value="0">Pyramid100PercentStacked</asp:ListItem>
								<asp:ListItem Value="1">PyramidBar100PercentStacked</asp:ListItem>
							</asp:DropDownList></td>
					</tr>
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
