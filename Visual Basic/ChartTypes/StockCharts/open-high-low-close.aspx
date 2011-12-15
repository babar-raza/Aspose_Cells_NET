<%@ Page Language="vb" Codebehind="open-high-low-close.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.OpenHighLowClose" MasterPageFile="~/tpl/Demo.Master"
	Title="Open-High-Low-Close - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tbody>
			<tr>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td  style="width: 100%;">
					<h2 class="demos-heading-bg">
						Open High Low Close - Aspose.Cells
					</h2>
				</td>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
			</tr>
		</tbody>
	</table>

<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
	<p>
	This online demo exhibits how to create an Open-High-Low-Close Stock chart in a worksheet using
   <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
							Aspose.Cells</a> for .NET.
	</p>

	<p>
	   The open-high-low-close chart is often used to illustrate stock prices. This type of chart requires four series of values 
	   in the correct order (open, high, low, and then close). Aspose.Cells is a powerful component, which supports all the 
	   standard and custom charts to help you display data in more meaningful ways. 
	</p>

	<p>
		The demo creates a workbook first and inputs some chart related data into the first five columns (A, B, C, D and E) 
		of the first worksheet named Data. The first column represents the companies (Microsoft, Mutual Fund 1 and Mutual Fund 2), 
		which denotes category data where as the second, third, fourth and fifth columns represent stock price values related to
		 the scenarios (Open, High, Low and Close). The demo creates an open-high-low-close stock chart representing Stock chart 
		 into the second worksheet named Chart based on the different stock price values of the four states (mentioned above) in 
		 the first worksheet. In the demo, you have been provided a sample snapshot of the chart and a command button labeled 
		 "Create Report" to create the chart. You are allowed to either open the resultant excel file into MS Excel or save 
		 directly to your disk.
	</p>
	<p>        
		Click <b>Create Report</b> to see how demo can  set the appearance properties of an open-high-low-close
		chart.
	</p>
</div>

	<table class="genericTable" style="text-align: left; font-family: Arial; font-size: small;"
		<tr>
			<td valign="top" align="right">
				<img alt="" src="../../Image/OpenHighLowClose.jpg"/></td>
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
