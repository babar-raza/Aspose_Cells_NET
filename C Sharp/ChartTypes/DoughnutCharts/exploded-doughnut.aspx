<%@ Page Language="c#" Codebehind="exploded-doughnut.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.ExplodedDoughnut" MasterPageFile="~/tpl/Demo.Master"
    Title="Exploded Doughnut - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Exploded Doughnut - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo describes how to create an <b>Exploded Doughnut Chart</b> in a
            worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            This chart type is like an exploded pie chart, but it can contain more than one
            data series. This type of chart displays the contribution of each value to a total
            while emphasizing individual values. Aspose.Cells is a powerful component, which
            supports all the standard and custom charts to help you display data in more meaningful
            ways. The demo creates a workbook first and inputs the chart source data into the
            first two columns (A and B) of the first worksheet named Data. The first column
            denotes the category data that represents the products (Apple and Orange ) where
            as the second column represents the yearly sales data that mentions values.
        </p>
        <p>
            The demo creates an exploded doughnut chart representing fruit sales by region for
            years into the second worksheet named Chart based on the different fruit sale values
            in the first worksheet. In the demo, you are provided a sample snapshot of the chart
            and a command button labeled Create Report to create the chart. You can either open
            the resultant excel file into <b>MS Excel</b> or save directly to your disk.
        </p>
        <p>
            Click <b>Create Report</b> to see how demo can &nbsp;set the appearance properties
            of a exploded doughnut chart.</p>
        <table class="genericTable" style="font-family: Arial; font-size: small;">
            <tr>
                <td valign="top" align="right">
                    <img alt="" src="../../Image/ExplodedDoughnut.jpg" /></td>
                <td valign="top" align="left">
                    <table class="genericTable">
                        <tr>
                            <td>
                                Save Format:
                            </td>
                            <td>
                                <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                                    <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                                    <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <asp:Button ID="btnProcess" runat="server" Text="Create Report"></asp:Button></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
</asp:Content>
