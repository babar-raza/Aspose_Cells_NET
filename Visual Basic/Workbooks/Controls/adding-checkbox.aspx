<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="adding-checkbox.aspx.vb" Inherits="Workbooks_Controls_AddCheckbox"
	Title="Adding CheckBox - Aspose.Cells Demos" %>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="Server">
	<table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
		style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
		<tbody>
			<tr>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td class="demos-heading-bg" style="width: 100%;">
					<h2 class="demos-heading-bg">
						Adding CheckBox - Aspose.Cells
					</h2>
				</td>
				<td style="width: 19; vertical-align: top;">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This demo shows how to <b>Add Checkbox</b> control in your worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			CheckBoxes are handy if you want to provide a way for a user to choose between two
			options, such as <b>true</b> or <b>false</b>; yes or no. Aspose.Cells allows you
			to use checkboxes in your worksheets, if desired. Aspose.Cells provides <b>CheckBoxes
				class</b>, which is used to add a new checkbox to the collection. There is another
			class CheckBox, which represents a checkbox used to define all types of settings.
			The demo creates an excel file. Then by using simple Aspose.Cells APIs it adds a
			checkbox and apply different setting to it.
		</p>
		<p>
			Click <b>Process </b>to see how example adds a CheckBox control to the workbook
			by setting different options. You can either open the resulting excel file into
			<b>MS Excel</b> or save directly to your disk.
		</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" /></p>
	</div>
</asp:Content>
