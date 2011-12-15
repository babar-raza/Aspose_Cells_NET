<%@ Page Language="vb" CodeBehind="managing-worksheets.aspx.vb" AutoEventWireup="false"
	Inherits="Aspose.Cells.Demos.ManagingWorksheets" MasterPageFile="~/tpl/Demo.Master"
	Title="Managing Worksheets - Add/Remove Worksheets - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Managing Worksheets - Add/Remove Worksheets - Aspose.Cells
					</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo describes how to <b>manipulate (Add, Remove) worksheets</b> in
			a workbook using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.</p>
		<p>
			A workbook is a multi-page excel document. Each page in it is called a worksheet.
			Aspose.Cells component is capable of manipulating worksheets. For example, you can
			insert, remove and rename a sheet in a workbook. The demo offers you two command
			buttons <b>Add</b> and <b>Remove</b> to exercise the manipulation tasks. When you
			click on <b>Add</b> button, a workbook is created with a default worksheet. It changes
			the name of the worksheet to <b>My Worksheet1</b>. It then, adds two more worksheets
			to it named <b>My Worksheet2</b> and <b>My Worksheet3</b> respectively. When you
			click on Remove button, the demo makes use of an existing excel file which has three
			worksheets in it. It then, removes the second worksheet from the worksheets collection
			in the workbook.You can either open the resulting excel file into <b>MS Excel </b>
			or save directly to your disk.
		</p>
		<p>
			Please download the
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/Workbooks/ManagingWorksheets.xls">ManagingWorksheets.xls</asp:HyperLink>
			used in this demo.</p>
		<p>
			<asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
				<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
				<asp:ListItem Value="XLSX">XLSX</asp:ListItem>
			</asp:DropDownList>
			<asp:Button runat="server" ID="Button1" Text=" Add " />&nbsp;
			<asp:Button runat="server" ID="Button2" Text="Remove" />
		</p>
	</div>
</asp:Content>
