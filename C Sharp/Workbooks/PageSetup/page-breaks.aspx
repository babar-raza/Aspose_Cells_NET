<%@ Page Language="c#" CodeBehind="page-breaks.aspx.cs" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.PageBreaks"
    MasterPageFile="~/tpl/Demo.Master" Title="Manage Page Breaks - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Manage Page Breaks - Aspose.Cells
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
            This online demo exhibits <b>how to set page breaks (horizontal and vertical)</b>
            in the worksheets using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            A <b>Page Break</b> is the place in your data file where one page ends and the next
            one begins. Aspose.Cells component allows you to add page breaks at any selected
            cell of the worksheet. The location of the cell where the page break is added, page
            is ended and the rest of the data after the page break will be printed on the next
            page while printing. Moreover you may remove any existing page break too.
        </p>
        <p>
            The demo creates an excel file and inputs some data into different cells of the
            first and second column in the first worksheet of the workbook. i.e., A1, A2, A3,
            B1, B2, B3. You are provided three command buttons for exercise. When you click
            Add button, a page break is added at B2 cell. When you click Remove button, some
            page breaks are added first and then removed using their index number in the collection.
            When you click Clear button, all the page breaks (horizontal and vertical) are deleted
            in the worksheet.
        </p>
        <p>
            You can either open the resulting excel file into MS Excel or save directly to your
            disk.
        </p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
        </p>
        <table>
            <tr>
                <td>
                    <asp:Button ID="Button1" Text=" Add " runat="server"></asp:Button>&nbsp;&nbsp;
                    <asp:Button ID="Button3" Text="Clear" runat="server"></asp:Button>&nbsp;&nbsp;
                    <asp:Button ID="Button2" Text="Remove" runat="server"></asp:Button>
                </td>
            </tr>
        </table>
    </div>
</asp:Content>
