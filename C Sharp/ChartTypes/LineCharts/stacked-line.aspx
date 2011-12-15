<%@ Page Language="c#" Codebehind="stacked-line.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.StackedLine" MasterPageFile="~/tpl/Demo.Master"
    Title="Stacked Line - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Stacked Line - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo demonstrates how to create a <b>Stacked Line Chart</b> in a worksheet
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            This type of chart displays the trend of the contribution of each value over time
            or categories. It is also available with markers displayed at each data value. Aspose.Cells
            component supports all the standard and custom charts including Stacked Line chart
            to help you display data in more meaningful ways. The component can create the chart
            into the worksheet in a workbook using the simplest APIs with ease. The demo creates
            a workbook first and inputs the source data related chart into the first six columns
            (A, B, C, D, E and F) of the first worksheet named Data. The first column represents
            different regions where as the second, third, fourth, fifth and sixth columns represent
            the sales data representing values involving different years (2002 - 2006).
        </p>
        <p>
            The demo creates a stacked line chart representing Sales By Region For Years into
            the second worksheet named Chart based on the different sales values of different
            regions in different years in the first worksheet. In the demo, you are provided
            a sample snapshot of the chart, a drop down list which represents whether you want
            to create the chart with data markers and a command button labeled Create Report
            to create and exercise the chart based on your selection from the drop down list.
            You can either open the resultant excel file into <b>MS Excel</b> or save directly
            to your disk.
        </p>
        <p>
            Click <b>Create Report</b> to see how demo can set the appearance properties of
            a stacked line chart.</p>
    </div>
    <table class="genericTable" style="font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" align="right">
                <img alt="" src="../../Image/StackedLine.jpg" /></td>
            <td valign="top" align="left">
                <table class="genericTable">
                    <tr>
                        <td>
                            Chart Type:</td>
                        <td>
                            <asp:DropDownList runat="server" ID="ChartTypeList">
                                <asp:ListItem Value="0">LineStacked</asp:ListItem>
                                <asp:ListItem Value="1">LineStackedWithDataMarkers</asp:ListItem>
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
