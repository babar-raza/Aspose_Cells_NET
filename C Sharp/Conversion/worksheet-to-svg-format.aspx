<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    Inherits="Aspose.Cells.Demos.Conversion.Worksheet2Svg" CodeBehind="worksheet-to-svg-format.aspx.cs" 
    Title="Convert Worksheet to SVG format - Aspose.Cells Demos"
    %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Convert Worksheet to SVG Format - Aspose.Cells</h2>
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
                Aspose.Cells</a> for .NET to convert a worksheet to svg format.
        </p>
        <p>
            The demo utilizes a template file which contains some sample data. When you click
            on the Process button, it executes a code which converts each worksheet of the source
            worbook into svg format. Finally, the demo shows the links of the output svg files
            which you can download and view in Web Browser e.g Internet Explorer or FireFox.
        </p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/ProductList.xls">ProductList.xls</asp:HyperLink>
            used in this demo.</p>
        <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
        <br />
        <br />
        <br />
        <asp:Panel ID="outPanel" runat="server">
        </asp:Panel>
    </div>
</asp:Content>
