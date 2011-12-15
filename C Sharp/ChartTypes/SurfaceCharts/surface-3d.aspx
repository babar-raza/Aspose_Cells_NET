<%@ Page Language="c#" Codebehind="surface-3d.aspx.cs" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.Surface3D"
    MasterPageFile="~/tpl/Demo.Master" Title="Surface3D - Aspose.Cells Demos" %>

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
                        Surface 3D - Aspose.Cells
                    </h2>
                </td>
                <td style="width: 19; vertical-align: top;">
                    <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo demonstrates how to create a 3-D Surface chart in a worksheet using
            <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.
        </p>
        <p>
            This type of chart shows trends in values across two dimensions in a continuous
            curve. For example, the following surface chart shows the various combinations of
            temperature and time that result in the same measure of tensile strength. The colors
            in this chart represent specific ranges of values. Displayed without color, a 3-D
            surface chart is called a wireframe 3-D surface chart. Aspose.Cells is a powerful
            component, which supports all the standard and custom charts to help you display
            data in more meaningful ways.
        </p>
        <p>
            The demo creates a workbook first and inputs some chart related data into the first
            six columns (A, B, C, D, E and F) of the first worksheet named 3D Surface. The first
            column denotes a category, time (in seconds), related to different ranges (0.2 -
            1.0) where as the second, third, fourth, fifth and sixth columns contain values
            of temperature series (10, 20, 30, 40 and 50). The demo creates a 3-D surface chart
            titled Tensile strength Measurements into the worksheet based on time and temperature
            values. In the demo, you are provided a sample snapshot of the chart, a drop down
            list that represents chart type (Surface3D and SurfaceWirframe3D) and a command
            button labeled Create Report to create and exercise the chart based on your selection
            from the drop down list. You may either open the resultant excel file into <b>MS Excel</b>
            or save directly to your disk.
        </p>
        <p>
            Click <b>Create Report</b> to see how demo can set the appearance properties of
            a 3D surface chart.
        </p>
    </div>
    <table class="genericTable" style="text-align: left; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" align="right">
                <img alt="" src="../../Image/Surface3D.jpg" /></td>
            <td valign="top" align="left">
                <table class="genericTable">
                <tr>
                        <td>
                            Chart Type:</td>
                        <td>
                            <asp:DropDownList runat="server" ID="ChartTypeList">
                                <asp:ListItem Value="0">Surface3D</asp:ListItem>
                                <asp:ListItem Value="1">SurfaceWireframe3D</asp:ListItem>
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
