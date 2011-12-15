<%@ Page Language="c#" Codebehind="column-3d.aspx.cs" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.Column3D"
    MasterPageFile="~/tpl/Demo.Master" Title="Column3D - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Column 3D - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo exhibits how to create a 3-D <b>Column Chart</b> with simple, clustered
            and stacked flavors in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            This type of chart compares data points along two axes. Aspose.Cells component is
            a powerful component, which supports all the standard and custom charts to help
            you display data in more meaningful ways. The demo creates a workbook first and
            inputs some chart related data into the first two columns (A and B) of the first
            worksheet. The first column represents the category data (Region) and the second
            column represents values (Marketing Costs).
        </p>
        <p>
            The demo creates a 3-D column chart representing marketing costs by region based
            on the different costs involving different regions. In the demo, you have been provided
            a sample snapshot of the chart and a few controls that represent the related list
            of chart data including a drop down which represents chart type, another five drop down lists
            which represent wall color, floor color, rotation angle, elevation angle and depth
            in percentage and a command button labeled Create Report to create and exercise
            the chart using your desired inputs. You can either open the resultant excel file
            into <b>MS Excel</b> or save directly to your disk to check the results.
        </p>
        <p>
            Click <b>Create Report</b> to see how demo can set the appearance properties of
            a 3D column chart.</p>
    </div>
    <table class="genericTable" style="font-family: Arial; font-size: small;">
        <tr>
            <td align="right">
                <img alt="" src="../../Image/Column3D.jpg" /></td>
            <td valign="top" align="left">
                <table class="genericTable">
                    <tr>
                        <td>
                            Chart Type:</td>
                        <td>
                            <asp:DropDownList ID="ColumnType" runat="server">
                                <asp:ListItem Value="0">Column3D</asp:ListItem>
                                <asp:ListItem Value="1">Column3DClustered</asp:ListItem>
                                <asp:ListItem Value="2">Column3DStacked</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            Walls Color:</td>
                        <td>
                            <asp:DropDownList ID="WallsColor" runat="server">
                                <asp:ListItem Value="0">Black</asp:ListItem>
                                <asp:ListItem Value="1" Selected="True">White</asp:ListItem>
                                <asp:ListItem Value="2">Red</asp:ListItem>
                                <asp:ListItem Value="3">Lime</asp:ListItem>
                                <asp:ListItem Value="4">Blue</asp:ListItem>
                                <asp:ListItem Value="5">Yellow</asp:ListItem>
                                <asp:ListItem Value="6">Magenta</asp:ListItem>
                                <asp:ListItem Value="7">Cyan</asp:ListItem>
                                <asp:ListItem Value="8">Maroon</asp:ListItem>
                                <asp:ListItem Value="9">Green</asp:ListItem>
                                <asp:ListItem Value="10">Navy</asp:ListItem>
                                <asp:ListItem Value="11">Olive</asp:ListItem>
                                <asp:ListItem Value="12">Purple</asp:ListItem>
                                <asp:ListItem Value="13">Teal</asp:ListItem>
                                <asp:ListItem Value="14">Silver</asp:ListItem>
                                <asp:ListItem Value="15">Gray</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            Floor Color:</td>
                        <td>
                            <asp:DropDownList ID="FloorColor" runat="server">
                                <asp:ListItem Value="0">Black</asp:ListItem>
                                <asp:ListItem Value="1" Selected="True">White</asp:ListItem>
                                <asp:ListItem Value="2">Red</asp:ListItem>
                                <asp:ListItem Value="3">Lime</asp:ListItem>
                                <asp:ListItem Value="4">Blue</asp:ListItem>
                                <asp:ListItem Value="5">Yellow</asp:ListItem>
                                <asp:ListItem Value="6">Magenta</asp:ListItem>
                                <asp:ListItem Value="7">Cyan</asp:ListItem>
                                <asp:ListItem Value="8">Maroon</asp:ListItem>
                                <asp:ListItem Value="9">Green</asp:ListItem>
                                <asp:ListItem Value="10">Navy</asp:ListItem>
                                <asp:ListItem Value="11">Olive</asp:ListItem>
                                <asp:ListItem Value="12">Purple</asp:ListItem>
                                <asp:ListItem Value="13">Teal</asp:ListItem>
                                <asp:ListItem Value="14">Silver</asp:ListItem>
                                <asp:ListItem Value="15">Gray</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            Rotation:</td>
                        <td>
                            <asp:DropDownList ID="Rotation" runat="server">
                                <asp:ListItem Value="0">0</asp:ListItem>
                                <asp:ListItem Value="20" Selected="True">20</asp:ListItem>
                                <asp:ListItem Value="40">40</asp:ListItem>
                                <asp:ListItem Value="60">60</asp:ListItem>
                                <asp:ListItem Value="80">80</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            Elevation:</td>
                        <td>
                            <asp:DropDownList ID="Elevation" runat="server">
                                <asp:ListItem Value="0" Selected="True">0</asp:ListItem>
                                <asp:ListItem Value="20">20</asp:ListItem>
                                <asp:ListItem Value="40">40</asp:ListItem>
                                <asp:ListItem Value="60">60</asp:ListItem>
                                <asp:ListItem Value="80">80</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            DepthPercent:</td>
                        <td>
                            <asp:DropDownList ID="DepthPercent" runat="server">
                                <asp:ListItem Value="0">100</asp:ListItem>
                                <asp:ListItem Value="1">200</asp:ListItem>
                                <asp:ListItem Value="2">300</asp:ListItem>
                                <asp:ListItem Value="3">400</asp:ListItem>
                                <asp:ListItem Value="4">500</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            Save Format:</td>
                        <td>
                            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td align="middle" colspan="2">
                            <asp:Button ID="btnProcess" runat="server" Text="Create Report" OnClick="btnProcess_Click">
                            </asp:Button></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</asp:Content>
