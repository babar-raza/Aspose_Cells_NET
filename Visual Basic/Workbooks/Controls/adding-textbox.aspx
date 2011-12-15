<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="adding-textbox.aspx.vb" Inherits="Workbooks_Controls_AddTextbox"
	Title="Adding TextBox - Aspose.Cells Demos" %>

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
						Adding TextBox - Aspose.Cells
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
			This demo shows how to <b>Add Textbox</b> control in your worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			One way to stress important information on your report is to use a textbox. For
			example, you can enter text to highlight your company's name or to indicate the
			geographic region with the highest sales etc. Aspose.Cells provides <b>TextBoxes class</b>,
			which is used to add a new text box to the collection. There is another class TextBox,
			which represents a text box used to define all types of settings. The demo creates
			an excel file. Then by using simple Aspose.Cells APIs it adds two textboxes and
			apply different format setting to them.
		</p>
		<p>
			Click <b>Process </b>to see how example adds two different TextBoxes in the workbook
			by setting different formatting options. You can either open the resulting excel
			file into <b>MS Excel</b> or save directly to your disk.
		</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
		</p>
	</div>
</asp:Content>
