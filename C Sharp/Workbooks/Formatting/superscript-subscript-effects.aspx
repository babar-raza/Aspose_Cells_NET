<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="superscript-subscript-effects.aspx.cs" Inherits="Workbooks_Formatting_SuperscriptSubscript"
    Title="Apply Superscript & Subscript Effect - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Apply Superscript & Subscript Effect - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo exhibits how to apply <b>Superscript</b> & <b>Subscript</b> effects
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            Aspose.Cells provides the functionality to set the Superscript and Subscript effects
            on <b>Font</b>. In this demo we will apply superscript and subscript effects on
            sample data to show how easy it is to implement this feature using Aspose.Cells.
            After pressing <b>Process</b> you can either open the resultant excel files into
            <b>MS Excel</b> or save directly to your disk to check the results.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
        </p>
    </div>
</asp:Content>
