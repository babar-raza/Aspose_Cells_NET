<%@ Page Language="vb" Codebehind="high-low-close.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.HighLowClose" MasterPageFile="~/tpl/Demo.Master"
	Title="High Low Close - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
   <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tbody>
			<tr>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td style="width: 100%;">
					<h2 class="demos-heading-bg">
						High Low Close - Aspose.Cells
					</h2>
				</td>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
			</tr>
		</tbody>
	</table>


<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
	<p>
		This online demo exhibits how to create a High-Low-Close <b>Stock chart</b> in a worksheet using
		<a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
						Aspose.Cells</a> for .NET.
	</p>

	<p>
	   The high-low-close chart is often used to illustrate stock prices. It requires three series of values in the following order 
	   (high, low, and then close). Aspose.Cells is a powerful component, which supports all the standard and custom charts to help 
	   you display data in more meaningful ways.
	</p>
	<p>
		The demo creates a workbook first and inputs some chart related data into the first four columns (A, B, C and D) of 
		the first worksheet named HighLowClose. The first column represents the companies (Microsoft, Mutual Fund 1 and Mutual Fund 2), 
		which denotes category data where as the second, third and fourth columns represent stock price values related to the scenarios 
		(High, Low and Close). The demo creates a high-low-close stock chart representing stock chart into the worksheet based on the 
		different stock price values of the three states mentioned above. In the demo, you have been provided a sample snapshot of the 
		chart and a command button labeled "<b>Create Report</b>" to create the chart. You can either open the resultant excel file into 
		<b>MS Excel</b> or save directly to your disk.
	</p>
	<p>
		Click <b>Create Report</b> to see how demo can  set the appearance properties of a high-low-close chart.
	</p>
   </div>
	<table class="genericTable" style="text-align: left; font-family: Arial; font-size: small;">
		<tr>
			<td valign="top" align="right">
				<img alt="" src="../../Image/HighLowClose.jpg"/></td>
			<td valign="top" align="left">
				<table>
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
