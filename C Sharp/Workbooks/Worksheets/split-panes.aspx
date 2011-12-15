<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="split-panes.aspx.cs" Inherits="Workbooks_Worksheets_SplitPanes" Title="Split Panes - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Split Panes - Aspose.Cells
                    </h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo describes how to <b>Split Panes</b> in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            If you need to split the screen to give you two different views into the same worksheet,
            use split panes. Excel offers a very handy feature to allow you to view more than
            one copy of your worksheet and for you to be able to scroll through each pane of
            your worksheet independently. You can do this by using <b>Split Panes</b> feature.When
            you split panes, the panes of your worksheet work simultaneously. If you make a
            change in one, it will simultaneously appear in the other. Aspose.Cells provides
            Split Panes feature for the users.
        </p>
        <p>
            Click <b>Execute</b> to see how example creates an excel file and use the split
            panes in the first worksheet. You can either open the resulting excel file into
            <b>MS Excel </b>or save directly to your disk.
        </p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/book1.xls">book1.xls</asp:HyperLink>
            used in this demo.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Execute" OnClick="btnExecute_Click" />
        </p>
    </div>
</asp:Content>
