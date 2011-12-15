<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="font-setting.aspx.cs" Inherits="Workbooks_Formatting_FontSetting"
    Title="Font Setting - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Font Settings - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo shows how to <b>Customize Fonts</b> and <b>Color</b> of the cells
            in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            One of the primary benefits of using Aspose.Cells component is the ease with which
            you can give your worksheet a <b>Professional Appearance</b>. The demo creates an
            excel file and performs customization of fonts with their attributes. It makes use
            of a constant string Aspose and <b>sets text color</b>, <b>Make text Bold</b>, <b>Italic</b>,
            <b>Underline</b> and <b>Strikeout</b>. It also puts together <b>Superscript</b>
            and <b>Subscript</b> attributes, changes the font and its <b>Size</b> as well. You
            can either open the resultant excel file into <b>MS Excel</b> or save directly to
            your disk.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
        </p>
    </div>
</asp:Content>
