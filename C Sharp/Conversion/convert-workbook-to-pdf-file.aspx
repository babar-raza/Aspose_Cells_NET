<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    Inherits="Xls2Pdf" Title="Convert Workbook to PDF File - Aspose.Cells Demos"
    CodeBehind="convert-workbook-to-pdf-file.aspx.cs" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Convert Workbook to Pdf File - Aspose.Cells</h2>
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
                Aspose.Cells</a> for .NET to convert a workbook to Pdf file.</p>
        <p>
            The demo utilizes a template file which contains some formatted data. It then converts
            the workbook to a Pdf file. You can either open the output file into your Pdf viewer
            or save directly to your disk.</p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/MyTestBook1.xls">MyTestBook1.xls</asp:HyperLink>
            used in this demo.</p>
        <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
    </div>
</asp:Content>
