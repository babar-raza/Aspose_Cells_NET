<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="adding-radiobutton.aspx.cs" Inherits="Workbooks_Controls_AddRadioButton"
    Title="Adding RadioButton - Aspose.Cells Demos" %>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="Server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tbody>
            <tr>
                <td style="width: 19; vertical-align: top;">
                    <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
                </td>
                <td class="demos-heading-bg" style="width: 100%;">
                    <h2 class="demos-heading-bg">
                        Adding Radio Button - Aspose.Cells
                    </h2>
                </td>
                <td style="width: 19; vertical-align: top;">
                    <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This demo shows how to <b>Add Radio Button</b> control in your worksheet using <a
                href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            A Radio Button or an option button is a control made of a round box O. The user
            makes his or her decision by selecting or clicking the round box. Actually, a radio
            button is usually, if not always accompanied by other radio buttons. Such radio
            buttons appear and behave as a group. The user decides which button is valid by
            selecting only one of them. When the user clicks one button, its round box fills
            with a (big) dot. When one button in the group is selected, the other round buttons
            of the (same) group are empty. The demo creates an excel file. Then by using simple
            Aspose.Cells APIs it adds three radio buttons as a group and apply different format
            setting to them.
        </p>
        <p>
            Click <b>Process </b>to see how example adds RadioButtons in the workbook with different
            formatting options. You can either open the resulting excel file into <b>MS Excel</b>
            or save directly to your disk.
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
