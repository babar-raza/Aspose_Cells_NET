<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="adding-image-hyperlink.aspx.cs" Inherits="Workbooks_DrawingObjects_AddingImageHyperlink"
    Title="Adding Image Hyperlink - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Adding Image Hyperlink - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo describes how to <b>Image Hyperlinks</b> in a worksheet. using
            <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            Aspose.Cells allows you to add hyperlinks for images to your spreadsheets at runtime.
            You can set / modify <b>Screen Tip</b> and address for the link for your need. You
            can add a simple picture and make it as hyperlink.</p>
        <p>
            Click <b>Process </b>to see how example creates an excel file, inserts an image
            into the first worksheet of the spreadsheet and creates a <b>Hyperlink</b>.
            <br />
            You can either open the resulting excel file into <b>MS Excel</b> or save directly
            to your disk.
        </p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/Image/School.jpg">School.jpg</asp:HyperLink>
            used in this demo.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
        </p>
    </div>
</asp:Content>
