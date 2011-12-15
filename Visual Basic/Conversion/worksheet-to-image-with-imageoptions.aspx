<%@ Page Language="vb" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
	Inherits="Sheet2ImageWithOptions" Title="Convert Worksheet to Image with Image Options - Aspose.Cells Demos"
	CodeBehind="worksheet-to-image-with-imageoptions.aspx.vb" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
	<table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
		<tbody>
			<tr>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
				</td>
				<td class="demos-heading-bg" style="width: 100%">
					<h2 class="demos-heading-bg">
						Convert Worksheet to Image with Image Options - Aspose.Cells</h2>
				</td>
				<td valign="top" style="width: 19px">
					<img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
				</td>
			</tr>
		</tbody>
	</table>
	<div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
		<p>
			This online demo demonstrates the ability of <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
				Aspose.Cells</a> for .NET to convert a worksheet to image file using different
			image options.</p>
		<p>
			The demo utilizes a template file which contains some simple data. Then different
			image options are applied like vertical and horizontal resolution of image, image
			format and its TiffCompression etc.The worksheet is then converted to an image.
			You can either open the output image file into your picture viewer or save directly
			to your disk.</p>
		<p>
			Please download the
			<asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/MyTestBook1.xls">MyTestBook1.xls</asp:HyperLink>
			used in this demo.</p>
		<asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
	</div>
</asp:Content>
