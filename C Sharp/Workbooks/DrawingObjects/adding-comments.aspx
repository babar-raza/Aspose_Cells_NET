<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="adding-comments.aspx.cs" Inherits="Workbooks_DrawingObjects_AddingComments"
    Title="Adding Comments - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Adding Comments - Aspose.Cells</h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo exhibits how to <b>Add Comments to the cells</b> in a worksheet
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            When you are creating or editing a worksheet especially with others, It's often
            valuable to be able to <b>Make Comments</b> that are not printed on the sheet but
            can be viewed on the screen. Aspose.Cells 's Comments feature fits the bill. You
            can insert comments, view comments and even edit the existing comments in the worksheet.
            You may also remove comments of the cells. The demo creates an excel file. It adds
            a comment to <b>B1</b> cell of the first worksheet, accesses the comment and gives
            it a string note. It also sets <b>font</b>, <b>height</b> and <b>width</b> of the
            comment. You can either open the resulting excel file into <b>MS Excel</b> or save
            directly to your disk.
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
