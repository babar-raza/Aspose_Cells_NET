<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	Inherits="AddingPdfBookmarks" Title="Adding Pdf Bookmarks - Aspose.Cells Demos"
	CodeBehind="adding-pdf-bookmarks.aspx.vb" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Adding Pdf Bookmarks - Aspose.Cells</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo demonstrates <b>how to add pdf bookmarks</b> using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET while converting a workbook to Pdf file.</p>
		<p>
			Aspose.Cells allows you to add bookmarks for your requirement at runtime. PDF bookmarks
			can improve the navigability of longer PDF documents. When adding bookmark links
			to other parts of a PDF document, you can have precise control over the exact view
			you want, you're not limited to just linking to a page or so. You set up the precise
			view by positioning the page as you would like it to be viewed, and then you create
			the bookmarks.
		</p>
		<p>
			You can either open the output file into your <b>Pdf Viewer</b> or save directly to your
			disk.</p>
		<asp:Button ID="Button1" runat="server" Text="Process" OnClick="btnExecute_Click" />
	</div>
</asp:Content>
