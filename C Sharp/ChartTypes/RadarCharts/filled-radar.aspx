<%@ Page Language="c#" Codebehind="filled-radar.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.FilledRadar" MasterPageFile="~/tpl/Demo.Master"
    Title="Filled Radar - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Filled Radar - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demodemonstrates how to create a <b>Filled Radar Chart</b> in worksheet
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            This type of chart displays changes in values relative to a center point. In this
            type of chart, the area covered by a data series is filled with a color. For example,
            in the following radar chart, the data series that covers the most area, Brand A,
            represents the brand with the highest vitamin content. Aspose.Cells is a powerful
            component, which supports all the standard and custom charts to help you display
            data in more meaningful ways.
        </p>
        <p>
            The demo creates a workbook first and inputs some chart related data into the first
            seven columns (A, B, C, D, E, F and G) of the first worksheet named Data. The first
            column represents the different brands (Brand A, Brand B and Brand C) where as the
            second, third, fourth, fifth, sixth and seventh columns represent percentage values
            related to vitamins (Vitamin A, Vitamin B1, Vitamin B2, Vitamin C, Vitamin D and
            Vitamin E). The demo creates a filled radar chart titled Nutritional Analysis into
            the second worksheet named Chart based on the different vitamin contents of different
            brands in the first worksheet. In the demo, you are provided a sample snapshot of
            the chart and a command button labeled <b>"Create Report"</b> to create and exercise
            the chart. You can either open the resultant excel file into MS Excel or save directly
            to your disk.</p>
        <p>
            Click <b>Create Report</b> to see how demo can set the appearance properties of
            a filled radar chart.</p>
    </div>
    <table class="genericTable" style="font-family: Arial; font-size: small; text-align: left;">
        <tr>
            <td valign="top" align="right">
                <img alt="" src="../../Image/FilledRadar.jpg" /></td>
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
