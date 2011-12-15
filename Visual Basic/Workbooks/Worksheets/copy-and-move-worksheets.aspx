<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	Inherits="Copy_Move" Title="Copy and Move Worksheets - Aspose.Cells Demos" CodeBehind="copy-and-move-worksheets.aspx.vb" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Copy and Move Worksheets - Aspose.Cells</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo explains how to implement <b>Copy</b> and <b>Move</b> operations
			using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET
		</p>
		<p>
			Aspose.Cells allows you to copy and move worksheets within a workbook, you may also
			copy and move worksheets between workbooks for your need. The demo uses a template
			excel file named <b>"Copy_Move.xls"</b>, which has three sheets in it i.e.., <b>"Copy"</b>,
			<b>"Move"</b> and <b>"Copy1"</b>. It, first, copies the first sheet's contents to
			the <b>"Copy1"</b> sheet, then, it moves the second sheet named <b>"Move"</b> to
			the last indexed position in the same workbook.</p>
		<p>
			Click <b>Process </b>to see how example implements copy and move operations using
			a template file. You can either open the output excel file into your MS Excel or
			save directly to your disk.
		</p>
		<p>
			Please download the
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/Workbooks/Copy_Move.xls">Copy_Move.xls</asp:HyperLink>
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
