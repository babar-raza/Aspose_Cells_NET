<%@ Page Language="c#" CodeBehind="image-as-background-fillformat-of-chart.aspx.cs"
    AutoEventWireup="True" Inherits="Aspose.Cells.Demos.ImageFillFormat" MasterPageFile="~/tpl/Demo.Master"
    Title="Setting Image as Background FillFormat of a Chart - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tbody>
            <tr>
                <td style="width: 19; vertical-align: top;">
                    <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
                </td>
                <td class="demos-heading-bg" style="width: 100%;">
                    <h2 class="demos-heading-bg">
                        Setting Image as Background FillFormat of a Chart - Aspose.Cells</h2>
                </td>
                <td style="width: 19; vertical-align: top;">
                    <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo exhibits how to set a <b>Image</b> as chart's background fillformat
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET</p>
        <p>
            The demo creates a simple pie chart and insert an image as the background fillformat
            of the chart. You can either open the resultant excel file into <b>MS Excel</b>
            or save directly to your disk to check the results.</p>
        <p>
            Click <b>Process</b> to see how demo can set the picture as background fillformat
            of a chart.</p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/Image/school.jpg">school.jpg</asp:HyperLink>
            used in this demo.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnProcess" runat="server" Text="Process" OnClick="btnProcess_Click" />
        </p>
    </div>
</asp:Content>
