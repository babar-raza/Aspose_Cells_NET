<%@ Page Language="vb" CodeBehind="smart-marker.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.SmartMarkerPage" MasterPageFile="~/tpl/Demo.Master"
	Title="Smart Marker - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Smart Marker - Aspose.Cells Demos</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			The demo illustrates how to create data reports with just a few lines of code based
			on <a href="http://www.aspose.com/documentation/.net-components/aspose.cells-for-.net/smart-markers.html">Smart
				Marker</a> technique.</p>
		<p>
			A <b>Smart Marker</b> represents a single data point or value that gives a means
			to <b>Aspose.Cells</b> to place relevant data into different cells of the worksheet
			in a workbook. We make use of <b>Designer</b> spreadsheets in which we write smart
			markers into different cells. Normally a smart marker consists of <b>DataSource</b>
			and a <b>Field Name</b> and starts with "&amp;=". The DataSource can be a <b>DataSet,
				DataTable, DataView</b> or an Object variable etc. You can also make use of
			dynamic formulas that allows you to insert MS Excel's formulas into cells even when
			the formula must reference rows that will be inserted during the export process.
			Moreover, you may calculate totals and sub totals of any data field too.</p>
		<p>
			The demo utilizes of designer spreadsheet excel file named <b>SmartMarkerDesigner.xls</b>
			stored in your application folder. The file has three sheets named "Customers",
			"Order Details" and "Variables". The smart markers written in the first two sheets
			(prefixed with their DataSources) represent different fields of the tables named
			"Customers" and "Order Details" of <b>Northwind</b> Database. The third sheet introduces
			smart markers based on some Object variables and the variables with parameters too.
			Click on "Designer spreadsheet with Smart Markers" hyperlink to open the file into
			MS Excel or save it to your disk. Click on "Get the result" hyperlink to open the
			resultant file with all the records and data into your browser or save the file
			to your disk.</p>
		<ul>
			<li>
				<p>
					<a href="smartmarker/designer.aspx">Designer spreadsheet with Smart Markers</a></p>
			</li>
			<li>
				<p>
					<a href="smartmarker/result.aspx">Get the result.</a></p>
			</li>
		</ul>
	</div>
</asp:Content>
