<%@ Page Language="c#" CodeBehind="setting-background-image-chartsheet.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.SettingBackgroundImageOfChartSheet" MasterPageFile="~/tpl/Demo.Master"
    Title="Setting Image as Background of Aspose.Cells.Charts.ChartType Worksheet - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Setting Image as Background of Aspose.Cells.Charts.ChartType Worksheet - Aspose.Cells
                    </h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo describes how to <b>set the background image</b> for a <b>Aspose.Cells.Charts.ChartType</b>
            worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            In this demo we will demonstrate how to set an image as background for a <b>Aspose.Cells.Charts.ChartType</b>
            worksheet. We will create a workbook with a <b>Aspose.Cells.Charts.ChartType</b>
            worksheet and then set an image as its background. Then we will create a simple
            chart on the worksheet to give a better understanding of how the background image
            will be displayed on a chart sheet. You can either open the resulting excel file
            into <b>MS Excel</b> or save directly to your disk.
        </p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/Image/school.JPG">school.jpg</asp:HyperLink>
            used in this demo.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button runat="server" ID="Button1" Text="Process" />
        </p>
    </div>
</asp:Content>
