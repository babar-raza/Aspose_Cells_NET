<%@ Page Language="vb" AutoEventWireup="true" CodeBehind="sales-totals-form.aspx.vb"
	Inherits="Aspose.Cells.Demos.Northwind.SalesTotalsForm" MasterPageFile="~/tpl/Demo.Master"
	Title="Sales Totals by Amount - Apose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Sales Totals by Amount - Aspose.Cells</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo illustrates how to create a well formatted Sales Totals by Amount
			report using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			Aspose.Cells component gives you the agility to report your data in a variety of
			ways. Aspose.Cells component is fully functional for creating all types of reports.
			You may customize the size and appearance of everything on a report. You can display
			the information the way you want to see it.
		</p>
		<p>
			The demo generates the report listing sales in descending order by amount. It Prints
			top 10 customers on first page and prints page total in page footer. ADO.NET is
			used to retrieve the data from the Orders and Customers table of Northwind database,
			to generate the report. You can either open the resultant excel file into MS Excel
			or save directly to your disk to check the results.</p>
		<p>
			Click Process to see how example Prints sales in descending order by amount, prints
			top 10 customers on first page, prints page total in page footer.
		</p>
		<p>
			Please download the
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/Northwind.xls">Northwind.xls</asp:HyperLink>
			used in this demo.</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" /></p>
	</div>
</asp:Content>
