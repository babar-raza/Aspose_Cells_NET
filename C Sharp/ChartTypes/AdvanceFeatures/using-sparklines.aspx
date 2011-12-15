<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="True"
    CodeBehind="using-sparklines.aspx.cs" Inherits="Aspose.Cells.Demos.UsingSparklines"
    Title="Using Sparklines - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td style="width: 19; vertical-align: top;">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
            </td>
            <td class="demos-heading-bg" style="width: 100%;">
                <h2 class="demos-heading-bg">
                    Using Sparklines - Aspose.Cells</h2>
            </td>
            <td style="width: 19; vertical-align: top;">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
            </td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This demo shows how to utilize sparklines feature using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            <b>Microsoft Excel 2010</b> can analyze information in more ways than ever before.
            It allows users to track and highlight important data trends with new data analysis
            and visualization tools. Sparklines are mini-charts that you could place inside
            cells so that you can view the data and the chart on the same table. With proper
            use of Sparklines, data analysis is quicker and more direct to the point. Developers
            can create add, delete or read sparklines (in the template file) for their need
            using the simplest APIs provided by Aspose.Cells. The Aspose.Cells.Charts namespace
            contains the APIs regarding Sparklines, so you need to import this namespace. Using
            the feature of adding custom graphics for a given data range, developers have the
            freedom to use different types of tiny charts to their desired cell areas.
        </p>
        <p>
            Click <b>Process</b> to see how demo can set the picture as background fillformat
            of a chart.</p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <%--<asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                --%><asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnProcess" runat="server" Text="Process" OnClick="btnProcess_Click" />
        </p>
    </div>
</asp:Content>
