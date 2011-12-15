<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	CodeBehind="extracting-ole-object.aspx.vb" Inherits="Workbooks_DrawingObjects_ExtractingOleObject"
	Title="Extract Ole Objects - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Extracting Ole Objects - Aspose.Cells</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;">
		<p class="componentDescriptionTxt" style="text-align: left">
			This online demo exhibits how to extract an ole object from a worksheet using <a
				href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET.
		</p>
		<p>
			<b>OLE</b> (Object Linking and Embedding) is Microsoft's framework for a compound
			document technology. Aspose.Cells supports to <b>add</b> / <b>manipulate Ole Objects</b>
			into your worksheets and can give more value to your workbooks. Aspose.Cells provides
			some important API related to the task. In this demo we will <b>Extract an Image</b>
			from an excel file inserted as an ole object in the worksheet.You can either open
			the resulting image file into an picture viewer or save directly to your disk.
		</p>
		<p>
			Please download the 
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/OleFile.xls">OleFile.xls</asp:HyperLink> used in this demo.
		</p>
		<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
	</div>
</asp:Content>
