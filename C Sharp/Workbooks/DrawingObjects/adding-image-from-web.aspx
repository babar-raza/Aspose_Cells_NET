<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="adding-image-from-web.aspx.cs" Inherits="Workbooks_DrawingObjects_AddingImageFromWeb"
    Title="Adding Image From Web - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Adding Image From Web - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo describes how to <b>Add Image from Web</b> into the worksheet in
            a workbook using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            Sometimes, you do require <b>inserting a picture from a URL</b> into an Excel file.
            You may do it quite easily. You just need to extract & download image data into
            stream and then you may use Aspose.Cells APIs to insert image (<b>from the stream</b>)
            into the worksheet.</p>
        <p>
            Click <b>Process </b>to see how example creates an excel file, inserts an image
            from the web into the first worksheet of the spreadsheet and returns the file to
            user.
            <br />
            You can either open the resulting excel file into <b>MS Excel</b> or save directly
            to your disk.
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
