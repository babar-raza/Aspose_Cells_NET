<%@ Page Language="c#" Codebehind="adding-textbox-in-chart.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.TextBoxInChart" MasterPageFile="~/tpl/Demo.Master"
    Title="Adding TextBox in Chart  - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td style="width: 19; vertical-align: top;">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="width: 100%;">
                <h2 class="demos-heading-bg">
                    Adding TextBox in Chart - Aspose.Cells</h2>
            </td>
            <td style="width: 19; vertical-align: top;">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo exhibits how to add a TextBox to an <b>Excel Chart</b> using <a
                href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET</p>
        <p>
            This demo demonstrates how to add a textbox control to a chart. The demo creates
            a simple Column chart and adds a textbox to the chart. You can either open the resultant
            excel file into <b>MS Excel</b> or save directly to your disk to check the results.</p>
        <p>
            Click <b>Process</b> to see how demo adds a textbox control to a chart.</p>
        <p class="componentDescriptionTxt">
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnProcess" runat="server" Text="Process" OnClick="btnProcess_Click" />&nbsp;</p>
    </div>
</asp:Content>
