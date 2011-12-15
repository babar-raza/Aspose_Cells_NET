<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="adding-wordart.aspx.cs" Inherits="Workbooks_DrawingObjects_AddingWordArt"
    Title="Adding WordArt - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Applying WordArt Styles to Text - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo describes how to apply <b>WordArt Styles</b> to text in a workbook
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            WordArt is a gallery of text styles that you can add to your worbook to create <b>decorative
                effects</b>, such as <b>Shadowed</b> or <b>Mirrored</b> (reflected) text. Aspose.Cells
            component gives you the ability to insert WordArt of differnt predefined text styles
            in your worksheet. You can specify the text, font name and size as well as the width
            and height of the text. In this demo we will add different WordArt Styles to show
            the feature.
        </p>
        <p>
            Click <b>Process </b>to see how example creates an excel file, inserts three different
            WordArt styles into the first worksheet of the workbook and returns the file to
            user. You can either open the resulting excel file into <b>MS Excel</b> or save
            directly to your disk.
        </p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <%--<asp:ListItem Value="XLSX">XLSX</asp:ListItem>--%>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
        </p>
    </div>
</asp:Content>
