<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="Result.aspx.vb" Inherits="Aspose.Cells.Demos.SmartMarker.Result"
	Title="Smart Markers Result - Aspose.Cells Demos" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeaderContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Smart Markers Result - Aspose.Cells</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo illustrates how to create data reports with just a few lines of
			code based on <a href="http://www.aspose.com/documentation/.net-components/aspose.cells-for-.net/smart-markers.html">
				Smart Marker</a> technique using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
					Aspose.Cells</a> for .NET.</p>
		<p>
			A Smart Marker represents a single data point or value that gives a mean to Aspose.Cells
			to place relevant data into different cells of the worksheet in a workbook. We make
			use of Designer spreadsheets in which we write smart markers into different cells.
			Normally a smart marker consists of DataSource and a Field Name and starts with
			&quot;&amp;=&quot;. The DataSource can be a DataSet, DataTable, DataView or an Object
			variable etc. You can also make use of dynamic formulas that allows you to insert
			MS Excel&#39;s formulas into cells even when the formula must reference rows that
			will be inserted during the export process. Moreover, you may calculate totals and
			sub totals of any data field too.</p>
		<p>
			The demo utilizes a designer spreadsheet excel file named SmartMarkerDesigner.xls
			stored in the application folder. The file has three sheets named &quot;Customers&quot;,
			&quot;Order Details&quot; and &quot;Variables&quot;. The smart markers written in
			the first two sheets (prefixed with their DataSources) represent different fields
			of the tables named &quot;Customers&quot; and &quot;Order Details&quot; of Northwind
			Database. The third sheet introduces smart markers based on some Object variables
			and the variables with parameters too.
		</p>
		<p>
			Click on <b>Process</b> button to open the resultant file with all the records and
			data into your browser or save the file to your disk.
		</p>
		<p>
			Please download the
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/SmartMarkerDesigner.xls">SmartMarkerDesigner.xls</asp:HyperLink>
			used in this demo.</p>
		<asp:Button ID="btnProcess" runat="server" Text="Process" OnClick="btnProcess_Click" />
	</div>
</asp:Content>
