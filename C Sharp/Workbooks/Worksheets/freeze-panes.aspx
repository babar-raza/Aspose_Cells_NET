<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="freeze-panes.aspx.cs" Inherits="Workbooks_Worksheets_FreezePanes"
    Title="Freeze Panes - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Freeze Panes - Aspose.Cells
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
            This online demo describes how to <b>freeze rows and columns (Freeze Panes)</b>
            in a workbook using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            While manipulating larger worksheets you may require freeze panes in your worksheet.
            Aspose.Cells component provides users with the capabilities like freezing rows &amp;
            columns at specified row / column index in the worksheet. Freezing panes helps you
            to view / work different areas of a larger worksheet at the same time. Freezing
            rows and columns allows you to select data that remains visible when scrolling in
            a worksheet. For example, keeping row and column labels visible as you scroll. The
            frozen rows are always the top rows while the frozen columns are always the left
            most columns. The demo creates an excel file and freezes the panes at C4 cell in
            the first worksheet in a workbook.</p>
        <p>
            Click <b>Execute</b> to see how example creates an excel file and freezes the panes
            at <b>C4</b> cell in the first worksheet in a workbook. You can either open the
            resulting excel file into <b>MS Excel</b> or save directly to your disk.
        </p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
        </p>
    </div>
</asp:Content>
