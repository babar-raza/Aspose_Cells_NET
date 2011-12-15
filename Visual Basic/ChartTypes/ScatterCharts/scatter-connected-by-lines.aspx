<%@ Page Language="vb" Codebehind="scatter-connected-by-lines.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.ScatterConnectedByLines" MasterPageFile="~/tpl/Demo.Master"
	Title="Scatter Connected By Lines - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
			<td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
				font-size: large;">
				<h2 class="demos-heading-bg">
					Scatter Connected By Lines - Aspose.Cells</h2>
			</td>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo exhibits how to create a <b>Scatter chart</b> connected by lines
			/ curves with or without data markers in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			This type of chart compares pairs of values. This type of chart can be displayed
			with or without straight or smoothed connecting lines between data points. These
			lines can be displayed with or without markers. The demo creates a scatter chart
			with connected lines / curves with or without data markers which shows uneven intervals
			(or clusters) of two sets of data. When you arrange your data for a scatter chart,
			place x values in one row or column, and then enter corresponding y values in the
			adjacent rows or columns.
		</p>
		<p>
			Aspose.Cells is a powerful component, which supports all the standard and custom
			charts to help you display data in more meaningful ways. The demo creates a workbook
			first and inputs chart related source data into the first two columns (A and B)
			of the first worksheet named Data. The first column provides Daily Rainfall that
			represents the x values where as the second column denotes Particulate that represents
			the y values. The demo creates a scatter chart representing Particulate Levels in
			Rainfall into second worksheet named Chart based on the x and y values in the first
			worksheet. In the demo, you are provided a sample snapshot of the chart, a drop
			down list which represents different types of scatter chart and a command button
			labeled Create Report to create and exercise the chart based on the type you selected.
			You are allowed to either open the resultant excel file into MS Excel or save directly
			to your disk.
		</p>
		<p>
			Click <b>Create Report</b> to see how demo can &nbsp;set the appearance properties
			of a scatter with data points connected by lines.</p>
	</div>
	<table class="genericTable" style="font-family: Arial; font-size: small;">
		<tr>
			<td align="right">
				<img alt="" src="../../Image/ScatterConnectedByLines.jpg" /></td>
			<td valign="top" align="left">
				<table class="genericTable">
					<tr>
						<td>
							Chart Type:</td>
						<td>
							<asp:DropDownList ID="ChartTypeList" runat="server">
								<asp:ListItem Value="0">ScatterConnectedByCurvesWithDataMarker</asp:ListItem>
								<asp:ListItem Value="1">ScatterConnectedByCurvesWithoutDataMarker</asp:ListItem>
								<asp:ListItem Value="2">ScatterConnectedByLinesWithDataMarker</asp:ListItem>
								<asp:ListItem Value="3">ScatterConnectedByLinesWithoutDataMarker</asp:ListItem>
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
