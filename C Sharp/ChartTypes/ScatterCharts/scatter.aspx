<%@ Page Language="c#" Codebehind="Scatter.aspx.cs" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.Scatter"
    MasterPageFile="~/tpl/Demo.Master" Title="Scatter - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td  style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Scatter - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo demonstrates how to create a Scatter chart in a worksheet using
            <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            This type of chart compares pairs of values. The demo creates a scatter chart which
            shows uneven intervals (or clusters) of two sets of data. When you arrange your
            data for a scatter chart, place x values in one row or column, and then enter corresponding
            y values in the adjacent rows or columns. Aspose.Cells is a powerful component,
            which supports all the standard and custom charts to help you display data in more
            meaningful ways.
        </p>
        <p>
            The demo creates a workbook first and inputs some chart related data into the first
            two columns (A and B) of the first worksheet named Scatter. The first column provides
            Daily Rainfall that represents the x values where as the second column denotes Particulate
            that represents the y values. The demo creates a scatter chart representing Particulate
            Levels in Rainfall into the worksheet based on the x and y values. In the demo,
            you are provided a sample snapshot of the chart and a command button labeled "Create
            Report" to create and exercise the chart. You can either open the resultant excel
            file into <b>MS Excel</b> or save directly to your disk.
        </p>
        <p>
            Click <b>Create Report</b> to see how demo can &nbsp;set the appearance properties
            of a scatter chart.</p>
    </div>
    <table class="genericTable" style="text-align: left; font-family: Arial; font-size: small;">
        <tr>
            <td align="right">
                <img alt="" src="../../Image/Scatter.jpg" /></td>
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
