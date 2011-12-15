<%@ Page Language="vb" Codebehind="Line.aspx.vb" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.Line"
	MasterPageFile="~/tpl/Demo.Master" Title="Line - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
			<td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
				font-size: large;">
				<h2 class="demos-heading-bg">
					Line Chart - Aspose.Cells</h2>
			</td>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo exhibits how to create a <b>Line Chart</b> in a worksheet using
			<a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			This type of chart displays trends over time or categories. It is also available
			with markers displayed at each data value. Aspose.Cells component supports all the
			standard and custom charts including Line chart to help you display data in more
			meaningful ways. The demo creates a workbook first and inputs the source data related
			chart into the first six columns (A, B, C, D, E and F) of the first worksheet named
			Line. The first column presents different regions where as the second, third, fourth,
			fifth and sixth columns represent the sales data representing values involving different
			years (2002 - 2006).
		</p>
		<p>
			The demo creates a Line chart representing Sales By Region For Years into the worksheet
			based on the different sales values of different regions in different years. In
			the demo, you are provided a sample snapshot of the chart, a few controls including
			five drop down lists which represent chart type (Line and LineWithDataMarker), marker
			style (Square, Triangle, Diamond, Circle, Dash, Dot, None etc.), marker background
			color, marker foreground color, marker size and a command button labeled Create
			Report to create and exercise the chart based on your selection from the drop down
			lists. You are allowed to either open the resultant excel file into <b>MS Excel</b>
			or save directly to your disk to check the results.
		</p>
		<p>
			Click <b>Create Report</b> to see how demo can set the appearance properties of
			a line chart.</p>
	</div>
	<table class="genericTable" style="font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" align="right">
				<img alt="" src="../../Image/Line.jpg" /></td>
			<td valign="top" align="left">
				<table class="genericTable">
					<tr>
						<td>
							Chart Type:</td>
						<td>
							<asp:DropDownList ID="ChartTypeList" runat="server">
								<asp:ListItem Value="0">Line</asp:ListItem>
								<asp:ListItem Value="1">LineWithDataMarkers</asp:ListItem>
							</asp:DropDownList></td>
					</tr>
					<tr>
						<td>
							NSeries MarkerStyle</td>
						<td>
							<asp:DropDownList runat="server" ID="NSeriesMarkerStyle">
								<asp:ListItem Value="0">Automatic</asp:ListItem>
								<asp:ListItem Value="1">Circle</asp:ListItem>
								<asp:ListItem Value="2">Dash</asp:ListItem>
								<asp:ListItem Value="3">Diamond</asp:ListItem>
								<asp:ListItem Value="4">Dot</asp:ListItem>
								<asp:ListItem Value="5">None</asp:ListItem>
								<asp:ListItem Value="6">Square</asp:ListItem>
								<asp:ListItem Value="7">SquarePlus</asp:ListItem>
								<asp:ListItem Value="8">SquareStar</asp:ListItem>
								<asp:ListItem Value="9" Selected="True">SquareX</asp:ListItem>
								<asp:ListItem Value="10">Triangle</asp:ListItem>
							</asp:DropDownList></td>
					</tr>
					<tr>
						<td>
							NSeries Marker BackColor:</td>
						<td>
							<asp:DropDownList runat="server" ID="NMarkBackColor">
								<asp:ListItem Value="0">Black</asp:ListItem>
								<asp:ListItem Value="1">White</asp:ListItem>
								<asp:ListItem Value="2">Red</asp:ListItem>
								<asp:ListItem Value="3">Lime</asp:ListItem>
								<asp:ListItem Value="4">Blue</asp:ListItem>
								<asp:ListItem Value="5">Yellow</asp:ListItem>
								<asp:ListItem Value="6">Magenta</asp:ListItem>
								<asp:ListItem Value="7">Cyan</asp:ListItem>
								<asp:ListItem Value="8">Maroon</asp:ListItem>
								<asp:ListItem Value="9">Green</asp:ListItem>
								<asp:ListItem Value="10">Navy</asp:ListItem>
								<asp:ListItem Value="11">Olive</asp:ListItem>
								<asp:ListItem Value="12">Purple</asp:ListItem>
								<asp:ListItem Value="13" Selected="True">Teal</asp:ListItem>
								<asp:ListItem Value="14">Silver</asp:ListItem>
								<asp:ListItem Value="15">Gray</asp:ListItem>
							</asp:DropDownList></td>
					</tr>
					<tr>
						<td>
							NSeries Marker ForeColor:</td>
						<td>
							<asp:DropDownList runat="server" ID="NMarkForeColor">
								<asp:ListItem Value="0">Black</asp:ListItem>
								<asp:ListItem Value="1">White</asp:ListItem>
								<asp:ListItem Value="2">Red</asp:ListItem>
								<asp:ListItem Value="3" Selected="True">Lime</asp:ListItem>
								<asp:ListItem Value="4">Blue</asp:ListItem>
								<asp:ListItem Value="5">Yellow</asp:ListItem>
								<asp:ListItem Value="6">Magenta</asp:ListItem>
								<asp:ListItem Value="7">Cyan</asp:ListItem>
								<asp:ListItem Value="8">Maroon</asp:ListItem>
								<asp:ListItem Value="9">Green</asp:ListItem>
								<asp:ListItem Value="10">Navy</asp:ListItem>
								<asp:ListItem Value="11">Olive</asp:ListItem>
								<asp:ListItem Value="12">Purple</asp:ListItem>
								<asp:ListItem Value="13">Teal</asp:ListItem>
								<asp:ListItem Value="14">Silver</asp:ListItem>
								<asp:ListItem Value="15">Gray</asp:ListItem>
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td>
							NSeries Marker Size</td>
						<td>
							<asp:DropDownList runat="server" ID="NSeriesMarkSize">
								<asp:ListItem Value="1">5</asp:ListItem>
								<asp:ListItem Value="2">8</asp:ListItem>
								<asp:ListItem Value="3">10</asp:ListItem>
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
