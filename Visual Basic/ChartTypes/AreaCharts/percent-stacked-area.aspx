<%@ Page Language="vb" Codebehind="percent-stacked-area.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.PercentStackedArea" MasterPageFile="~/tpl/Demo.Master"
	Title="Percent Stacked Area - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tr>
			<td style="width: 19; vertical-align: top;">
				<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
			<td class="demos-heading-bg" style="width: 100%;">
				<h2 class="demos-heading-bg">
					Percent Stacked Area - Aspose.Cells</h2>
			</td>
			<td style="width: 19; vertical-align: top;">
				<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
		</tr>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo demonstrates how to create a 100% <b>Stacked Area chart</b> with
			simple and 3-D visual effects in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			This chart type displays the trend of the percentage each value contributes over
			time or categories. It is also available with a 3-D visual effect. Aspose.Cells
			is a powerful component, which supports all the standard and custom charts to help
			you display data in more meaningful ways. The demo creates a workbook first and
			inputs some chart related data into the first five columns (A, B, C, D and E) of
			the first worksheet named Data. The first column represents the product names (Product1,
			Product2 and Product3) where as the second, third, fourth and fifth columns represent
			percentage values involving different quarters (Qurarter1, Quarter2, Quarter3 and
			Quarter4) which represent category data.
		</p>
		<p>
			The demo creates a 100% stacked area chart representing product contribution to
			total sales into the second worksheet named Chart based on the different product
			values related to different quarters in the first worksheet. In the demo, you are
			provided a sample snapshot of the chart, a check box that represents whether you
			want to create the chart with 3-D visual effect and a command button labeled Create
			Report to create and exercise the chart using your desired inputs. You can either
			open the resultant excel file into <b>MS Excel</b> or save directly to your disk
			to check the results.</p>
		<p>
			Click <b>Create Report</b> to see how demo can set the appearance properties of
			a 100% stacked area chart.</p>
	</div>
	<table class="genericTable" style="font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" align="right">
				<img alt="" src="../../Image/PercentStackedArea.jpg" /></td>
			<td valign="top" align="left">
				<table class="genericTable">
					<tr>
						<td>
							Show as 3D:</td>
						<td>
							<asp:CheckBox runat="server" ID="CheckBoxShow3D"></asp:CheckBox></td>
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
							<asp:Button runat="server" ID="btnProcess" Text="Create Report" OnClick="btnProcess_Click">
							</asp:Button></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</asp:Content>
