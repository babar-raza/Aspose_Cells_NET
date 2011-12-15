<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    Codebehind="adding-combobox.aspx.cs" Inherits="Workbooks_Controls_AddCombobox"
    Title="Adding ComboBox - Aspose.Cells Demos" %>

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
                        Adding ComboBox - Aspose.Cells
                    </h2>
                </td>
                <td style="width: 19; vertical-align: top;">
                    <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This demo shows how to <b>Add Combobox</b> control in your worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            To make data entry easier, or to limit entries to certain items that you define,
            you can create a <b>ComboBox</b> or <b>DropdownList</b> of valid entries that is compiled from
            cells elsewhere on the worksheet. The demo creates an excel file and by using simple
            Aspose.Cells APIs it adds a combobox by using range of Cells from <b>A2</b> to <b>A7</b> as data
            source.<br />
        </p>
        <p>
            Click <b>Process </b>to see how example adds a ComboBox control to the worksheet
            by using cell range as data source. You can either open the resulting excel file
            into <b>MS Excel</b> or save directly to your disk.
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
