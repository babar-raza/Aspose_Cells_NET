<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    Inherits="Sheet2ImageWithPrintArea" Title="Convert Worksheet to Image with multiple Print Areas - Aspose.Cells Demos"
    CodeBehind="convert-worksheet-to-image-with-multiple-printarea.aspx.cs" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Convert Worksheet to Image with multiple Print Areas - Aspose.Cells</h2>
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
                Aspose.Cells</a> for .NET to convert a worksheet to image file having multiple
            print areas.</p>
        <p>
            The demo converts the worksheet having multiple print areas defined in to an image
            file. You can either open the output image file into your picture viewer or save
            directly to your disk.</p>
        <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" /></div>
</asp:Content>
