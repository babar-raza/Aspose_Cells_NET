<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    Inherits="Workbook2Image" Title="Convert Workbook to Image - Aspose.Cells Demos"
    CodeBehind="convert-workbook-to-image.aspx.cs" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Convert Workbook to Image - Aspose.Cells</h2>
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
                Aspose.Cells</a> for .NET to convert a workbook to an image file.</p>
        <p>
            The demo converts the complete workbook (all worksheets) in to an image file. You
            can either open the output image file into your picture viewer or save directly
            to your disk.</p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/FinancialPlan.xls">FinancialPlan.xls</asp:HyperLink>
            used in this demo.</p>
        <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" /></div>
</asp:Content>
