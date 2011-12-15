<%@ Page Language="c#" Codebehind="Radar.aspx.cs" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.Radar"
    MasterPageFile="~/tpl/Demo.Master" Title="Radar - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Radar - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo demonstrates how to create a <b>Radar chart</b> in a worksheet using <a
                href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            This type of chart displays changes in values relative to a center point. It can
            be displayed with markers for each data point. For example, in the following radar
            chart, the data series that covers the most area, Brand A, represents the brand
            with the highest vitamin content. Aspose.Cells is a powerful component, which supports
            all the standard and custom charts to help you display data in more meaningful ways.
        </p>
        <p>
            The demo creates a workbook first and inputs some chart related data into the first
            seven columns (A, B, C, D, E, F and G) of the first worksheet named Radar. The first
            column represents the different brands (Brand A, Brand B and Brand C) where as the
            second, third, fourth, fifth, sixth and seventh columns represent percentage values
            related to vitamins (Vitamin A, Vitamin B1, Vitamin B2, Vitamin C, Vitamin D and
            Vitamin E). The demo creates a radar chart titled "Nutritional Analysis" into the
            worksheet based on the different vitamin content values of different brands. In
            the demo, you are provided a sample snapshot of the chart, a drop down list that
            represents the chart type (Radar, RadarWithDataMarkers) and a command button labeled
            <b>Create Report</b> to create and exercise the chart based on your selection from the
            drop down list. You can either open the resultant excel file into <b>MS Excel</b> or save
            directly to your disk to check the results.</p>
        <p>
            Click <b>Create Report</b> to see how demo can set the appearance properties of
            a radar chart.</p>
    </div>
    <table class="genericTable" style="font-family: Arial; font-size: small; text-align: left;">
        <tr>
            <td valign="top" align="right">
                <img alt="" src="../../Image/RadarUnfilled.jpg" /></td>
            <td valign="top" align="left">
                <table class="genericTable">
                    <tr>
                        <td>
                            Chart Type:</td>
                        <td>
                            <asp:DropDownList runat="server" ID="ChartTypeList">
                                <asp:ListItem Value="0">Radar</asp:ListItem>
                                <asp:ListItem Value="1">RadarWithDataMarkers</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
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
