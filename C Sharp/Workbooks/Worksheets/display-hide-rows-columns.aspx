<%@ Page Language="c#" CodeBehind="display-hide-rows-columns.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.DisplayHideRowsColumns" MasterPageFile="~/tpl/Demo.Master"
    Title="Display and Hide Row and Column Headers - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Display and Hide Row and Column Headers - Aspose.Cells
                    </h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo describes how to <b>display/hide</b> the <b>row/column</b> headers
            of a worksheet in the workbook using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p class="componentDescriptionTxt">
            Sometimes, you require to hide row/column headers to further improve readability
            in the worksheets. Aspose.Cells component allows you to control the visibility of
            the worksheet row/column headers. The demo offers you two command buttons <b>Hide</b>
            and <b>Display</b> to exercise the tasks. When you click on Hide button, a workbook
            is created. It makes the row / column headers invisible of the first worksheet in
            the workbook. When you click on Display button, the demo creates an excel file and
            makes the row / column headers visible of the first worksheet in the workbook. By
            default worksheet row / column headers are visible in the workbook. You can either
            open the resulting excel file into <b>MS Excel</b> or save directly to your disk.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="Button1" Text="Display" runat="server"></asp:Button>&nbsp;&nbsp;
            <asp:Button ID="Button2" Text="Hide" runat="server"></asp:Button>
        </p>
    </div>
</asp:Content>
