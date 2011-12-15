<%@ Page AutoEventWireup="false" Codebehind="Doughnut.aspx.cs" Inherits="Aspose.Cells.Demos.Doughnut"
    Language="c#" MasterPageFile="~/tpl/Demo.Master" Title="Doughnut - Aspose.Cells Demos"%>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Doughnut - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo describes how to create a <b>Doughnut Chart</b> in a worksheet
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            This type of chart displays data in rings, where each ring represents a data series.
            For example, in the following chart, the inner ring represents 2005 fruit sales,
            and the outer ring represents 2006 fruit sales. Aspose.Cells is a powerful component,
            which supports all the standard and custom charts to help you display data in more
            meaningful ways. The demo creates a workbook first and inputs the chart source data
            into the first three columns (A, B and C) of the first worksheet named Doughnut.
            The first column represents the products (Apple and Orange ) where as the second
            and third columns represent the sales data involving different years representing
            values.
        </p>
        <p>
            The demo creates a doughnut chart representing fruit sales by region for years into
            the worksheet based on the different fruit sale values related to different years.
            In the demo, you are provided a sample snapshot of the chart and a command button
            labeled Create Report to create the chart. You can either open the resultant excel
            file into <b>MS Excel</b> or save directly to your disk.
        </p>
        <p>
            Click <b>Create Report</b> to see how demo can set the appearance properties of
            a doughnut chart.
        </p>
    </div>
    <table class="genericTable" style="font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" align="right">
                <img alt="" src="../../Image/Doughnut.jpg" /></td>
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
