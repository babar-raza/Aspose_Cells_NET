<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="modify-existing-style.aspx.cs" Inherits="Workbooks_Formatting_ModifyExistingStyle"
    Title="Modify Existing Style - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%">
                <h2 class="demos-heading-bg">
                    Modify Existing Style - Aspose.Cells</h2>
            </td>
            <td valign="top" style="width: 19px">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo exhibits how to modify <b>Existing Style</b> using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            Aspose.Cells provides <b>Style.Update</b> method to update an existing style. For
            a named style (whether it is created dynamically using Aspose.Cells API or it belongs
            to Predefined styles list in <b>MS Excel</b>), if you want to change the Style,
            you may call Style.Update method or otherwise the <b>style of cell</b> / <b>cells range</b>
            (to whom you have applied formattings) will not be reflected. The Style.Update method
            behaves like an <b>OK button of the Style dialog box</b> i.e., when you have done
            your modifications with an existing style, you may call it for the final implementation.
            So, if you have already applied some style to a range of cell(s), then you modify
            the style attributes and finally call the method, the style formattings of those
            cells would be upldated too. You can either open the resultant excel files into
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
