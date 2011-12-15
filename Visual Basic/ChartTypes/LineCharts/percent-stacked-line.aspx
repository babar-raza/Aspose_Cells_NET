<%@ Page Language="vb" Codebehind="percent-stacked-line.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.PercentStackedLine" MasterPageFile="~/tpl/Demo.Master"
	Title="Percent Stacked Line - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
			<td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
				font-size: large;">
				<h2 class="demos-heading-bg">
					Percent Stacked Line - Aspose.Cells</h2>
			</td>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo demonstrates how to create a <b>100% Stacked Line Chart</b> in
			a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			This type of chart displays the trend of the percentage each value contributes over
			time or categories. It is also available with markers displayed at each data value.
			Aspose.Cells is a powerful component, which supports all the standard and custom
			charts to help you display data in more meaningful ways. The demo creates a workbook
			first and inputs some chart related data into the first five columns (A, B, C, D
			and E) of the first worksheet named Data. The first column represents the product
			names (Product1, Product2 and Product3) where as the second, third, fourth and fifth
			columns represent percentage values involving different quarters (Qurarter1, Quarter2,
			Quarter3 and Quarter4).
		</p>
		<p>
			The demo creates a 100% stacked line chart representing Product contribution to
			total sales into the second worksheet named Chart based on the different product
			values related to different quarters in the first worksheet. In the demo, you have
			been provided a sample snapshot of the chart, a drop down list that represents whether
			you want to create the chart with data markers and a command button labeled Create
			Report to create and exercise the chart based on the selection from the drop down
			list. You can either open the resultant excel file into <b>MS Excel</b> or save
			directly to your disk to check the results.
		</p>
		<p>
			Click <b>Create Report</b> to see how demo can set the appearance properties of
			a 100% stacked line chart.
		</p>
	</div>
	<table class="genericTable" style="font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" align="right">
				<img src="../../Image/PercentStackedLine.jpg"></td>
			<td valign="top" align="left">
				<table class="genericTable">
					<tr>
						<td>
							Chart Type:</td>
						<td>
							<asp:DropDownList runat="server" ID="ChartTypeList">
								<asp:ListItem Value="0">Line100PercentStacked</asp:ListItem>
								<asp:ListItem Value="1">Line100PercentStackedWithDataMarkers</asp:ListItem>
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
