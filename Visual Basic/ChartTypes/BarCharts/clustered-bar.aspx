<%@ Page Language="vb" Codebehind="clustered-bar.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.ClusteredBar" MasterPageFile="~/tpl/Demo.Master"
	Title="Clustered Bar - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tr>
			<td style="width: 19; vertical-align: top;">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
			<td class="demos-heading-bg" style="width: 100%;">
				<h2 class="demos-heading-bg">
					Clustered Bar - Aspose.Cells</h2>
			</td>
			<td style="width: 19; vertical-align: top;">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo exhibits how to create a <b>Clustered Bar chart</b> with 2-D and 3-D visual
			effects in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			This type of chart compares values across categories. Normally, categories are organized
			vertically, and values horizontally, to place focus on comparing the values. Aspose.Cells
			is a powerful component, which supports all the standard and custom charts to help
			you display data in more meaningful ways. The demo creates a workbook first and
			inputs the source data related chart into the first three columns (A, B and C) of
			the first worksheet. The first column represents the category data (Region) where
			as the second and third columns represent the sales data representing values related
			to the products (Apple and Orange).</p>
		<p>
			The demo creates a clustered bar chart representing fruit sales by region into the
			first worksheet named Clustered Bar based on the different product values related
			to different regions. In the demo, you are provided a sample snapshot of the chart,
			a check box that represents whether you want to create the chart with 3-D flavor
			and a command button labeled Create Report to create and exercise the chart using
			your desired inputs. You can either open the resultant excel file into <b>MS Excel</b>
			or save directly to your disk to check the results.</p>
		<p>
			Click <b>Create Report</b> to see how demo can set the appearance properties of
			a clustered bar chart.</p>
	</div>
	<table class="genericTable" style="font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" align="right">
				<img alt="" src="../../Image/ClusteredBar.jpg" /></td>
			<td valign="top" align="left">
				<table class="genericTable">
					<tr>
						<td>
							Show as 3D:</td>
						<td>
							<asp:CheckBox ID="checkBoxShow3D" runat="server"></asp:CheckBox></td>
					</tr>
					<tr>
						<td>
							Save Format:</td>
						<td>
							<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
								<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
								<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
							</asp:DropDownList></td>
					</tr>
					<tr>
						<td colspan="2">
							<asp:Button ID="btnProcess" runat="server" Text="Create Report" OnClick="btnProcess_Click">
							</asp:Button></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</asp:Content>
