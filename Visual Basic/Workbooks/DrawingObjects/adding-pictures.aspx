<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="adding-pictures.aspx.vb" Inherits="Workbooks_DrawingObjects_AddingPictures"
	Title="Adding Pictures - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Adding Pictures - Aspose.Cells</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;">
		<p>
			This online demo describes how to <b>Add Pictures</b> into the worksheet in a workbook
			using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			Sometimes, you may need to insert some pictures of your products or your <b>company
				logo</b> in the worksheet for more readability. Aspose.Cells component gives
			you the ability to insert pictures of <b>various formats</b> and enhance your worksheet.
			You can change the <b>height</b> and <b>width</b> of a picture while inserting into
			the worksheet. You can also put in a picture from stream. You may remove the pictures
			of the cells in your worksheet too.
		</p>
		<p>
			Click <b>Process </b>to see how example creates an excel file, inserts two pictures
			into the first worksheet of the spreadsheet and returns the file to user.
			<br />
			You can either open the resulting excel file into <b>MS Excel</b> or save directly
			to your disk.
		</p>
		<p>
			Please download the
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/Image/School.jpg">School.jpg</asp:HyperLink>
			used in this demo.</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
		</p>
	</div>
</asp:Content>
