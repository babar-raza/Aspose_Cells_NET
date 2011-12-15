<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="format-list-object.aspx.vb" Inherits="Aspose.Cells.Demos.Format_List_Object"
	Title="Formatting a List Object - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tbody>
			<tr>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td class="demos-heading-bg" style="width: 100%;">
					<h2 class="demos-heading-bg">
						Formatting a List Object - Aspose.Cells</h2>
				</td>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo exhibits how to <b>Set the Pppearance of a List Object</b> in a
			worksheet using<a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			To make managing and analyzing a group of related data easier, you can turn a <b>range
				of cells</b> into a <b>List Object</b> (also known as <b>Excel table</b>). A
			table is a series of rows and columns that contains related data that is managed
			independently from the data in other rows and columns on the worksheet. By default,
			every column in the table has <b>Filtering Enabled</b> in the header row so that
			you can filter or sort your List Object data quickly. You can add a total row (total
			row: A special row in a list that provides a selection of aggregate functions useful
			for working with numerical data.) to your list object that provides a <b>DropdownList</b>
			of aggregate functions for each total row cell. In this demo, we will add some sample
			data in the worksheet, add a List Object and apply some default style to the List
			Object. The List Objects Styles are supported in MS Excel 2007 file formats.</p>
		<p>
			Click <b>Process </b>to see how demo creates a List Object and set its default style.
			You can either open the resulting excel file into <b>MS Excel</b> or save directly
			to your disk.
		</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" /></p>
	</div>
</asp:Content>
