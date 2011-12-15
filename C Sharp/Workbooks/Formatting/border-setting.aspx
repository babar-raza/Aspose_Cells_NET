<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="border-setting.aspx.cs" Inherits="Workbooks_Formatting_BorderSetting"
    Title="Border Setting - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Border Settings - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo explains how to <b>Set Borders</b> to the cells in a worksheet
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            Borders and color provides you further ways to highlight information in the cells
            in a worksheet. The demo creates an excel file. It sets different borders (<b>Top</b>,
            <b>Bottom</b>, <b>Left</b>, <b>Right</b>, <b>DiagonalDown</b>, <b>DiagonalUp</b>)
            to <b>B2</b> cell of the first worksheet with a <b>DashDot line style</b>. It applies
            blue border color lines. You can either open the resultant excel file into <b>MS Excel</b>
            or save directly to your disk.
        </p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" /></p>
    </div>
</asp:Content>
