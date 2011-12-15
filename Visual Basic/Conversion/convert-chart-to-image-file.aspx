<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	Inherits="Chart2Image" Title="Convert Chart to Image File - Aspose.Cells Demos"
	CodeBehind="convert-chart-to-image-file.aspx.vb" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Convert Chart to Image File - Aspose.Cells</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo demonstrates how to convert a Pie chart to image file using 
			<a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">Aspose.Cells</a>
			for .NET API.</p>
		<p>
			The demo creates a workbook first and inputs the chart source data into the first
			two columns (A and B) of the first worksheet. The first column represents the category
			data (Region) where as the second column represents the sales data representing
			values.
		</p>
		<p>
			The demo creates a pie chart dynamically representing Sales By Region into the worksheet
			named ChartSheet based on the different sale values related to different regions.
			Then, it converts this chart in the first worksheet to image file. You can either
			open the resultant image file into your picture viewer or save directly to your
			disk.</p>
		<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
	</div>
</asp:Content>
