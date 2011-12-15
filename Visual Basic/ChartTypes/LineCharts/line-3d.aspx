<%@ Page AutoEventWireup="false" Codebehind="line-3d.aspx.vb" Inherits="Aspose.Cells.Demos.Line3D"
	Language="vb" MasterPageFile="~/tpl/Demo.Master" Title="Line 3D - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
			<td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
				font-size: large;">
				<h2 class="demos-heading-bg">
					Line 3D - Aspose.Cells</h2>
			</td>
			<td valign="top" style="height: 41; width: 19">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo exhibits how to create a <b>Line chart</b> with a <b>3-D</b> visual
			effect in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			Aspose.Cells component supports all the standard and custom charts including 3-D
			Line chart to help you display data in more meaningful ways. The demo creates a
			workbook first and inputs the chart source data into the first six columns (A, B,
			C, D, E and F) of the first worksheet named 3D Line. The first column presents different
			regions where as the second, third, fourth, fifth and sixth columns represent the
			sales data representing values involving different years (2002 - 2006).
		</p>
		<p>
			The demo creates a 3-D Line chart representing Sales By Region into the worksheet
			based on the different sales values of different regions in different years. In
			the demo, you are provided a sample snapshot of the chart, a few controls including
			four drop down lists which represent major tick mark type (None, Inside, Outside
			and Cross) and minor tick mark type (None, Inside, Outside and Cross) for values,
			value labels rotation angle and category labels rotation angle and a command button
			labeled Create Report to create and exercise the chart based on your selection from
			the drop down lists. You are allowed to either open the resultant excel file into
			<b>MS Excel</b> or save directly to your disk to check the results.
		</p>
		<p>
			Click <b>Create Report</b> to see how demo can set the appearance properties of
			a 3D line chart.</p>
	</div>
	<table class="genericTable" style="font-family: Arial; font-size: small;">
		<tr>
			<td align="right">
				<img alt="" src="../../Image/Line3D.jpg" /></td>
			<td valign="top" align="left">
				<table class="genericTable">
					<tr>
						<td>
							ValueAxis Major Tick Mark Type:</td>
						<td>
							<asp:DropDownList ID="MajorTickMarkType" runat="server">
								<asp:ListItem Value="0">None</asp:ListItem>
								<asp:ListItem Value="1">Inside</asp:ListItem>
								<asp:ListItem Value="2">Outside</asp:ListItem>
								<asp:ListItem Value="3">Cross</asp:ListItem>
							</asp:DropDownList></td>
					</tr>
					<tr>
						<td>
							ValueAxis Minor Tick Mark Type:</td>
						<td>
							<asp:DropDownList ID="MinorTickMarkType" runat="server">
								<asp:ListItem Value="0">None</asp:ListItem>
								<asp:ListItem Value="1">Inside</asp:ListItem>
								<asp:ListItem Value="2">Outside</asp:ListItem>
								<asp:ListItem Value="3">Cross</asp:ListItem>
							</asp:DropDownList></td>
					</tr>
					<tr>
						<td>
							ValueAxis TickLabels Rotation:</td>
						<td>
							<asp:DropDownList ID="VLabelsRotation" runat="server">
								<asp:ListItem Value="0">0</asp:ListItem>
								<asp:ListItem Value="1">5</asp:ListItem>
								<asp:ListItem Value="2">10</asp:ListItem>
							</asp:DropDownList></td>
					</tr>
					<tr>
						<td>
							CategoryAxis Ticklabels Rotation:</td>
						<td>
							<asp:DropDownList ID="CLabelsRotation" runat="server">
								<asp:ListItem Value="0">0</asp:ListItem>
								<asp:ListItem Value="1">5</asp:ListItem>
								<asp:ListItem Value="2">10</asp:ListItem>
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
