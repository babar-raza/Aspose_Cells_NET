<%@ Page Language="c#" Codebehind="clustered-column.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.ClusteredColumn" MasterPageFile="~/tpl/Demo.Master"
    Title="Clustered Column - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Clustered Column - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo exhibits how to create a <b>Clustered Column Chart</b> in a worksheet
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            This type of chart compares values across categories. Normally the categories are
            organized horizontally, and values vertically, to emphasize variation over time.
            Aspose.Cells component is a powerful component, which supports all the standard
            and custom charts to help you display data in more meaningful ways. The demo creates
            a workbook first and inputs some chart related data into the first two columns (A
            and B) of the first worksheet. The first column represents the category data (Region)
            and the second column represents values (Marketing Costs).
        </p>
        <p>
            The demo creates the chart representing marketing costs by region based on the different
            costs involving different regions. In the demo, you have been provided a sample
            snapshot of the chart and a few controls that represent the related list of data
            including two text boxes which represent category axis and value axis titles, five
            drop down lists which represent major unit, minor unit, minimum, maximum values
            of value axis and gap width that represents the space b/w the column clusters in
            the chart and a command button labeled Create Report to create and exercise the
            chart using your desired inputs. You can either open the resultant excel file into
            <b>MS Excel</b> or save directly to your disk to check the results.
        </p>
        <p>
            Click <b>Create Report</b> to see how demo can &nbsp;set the appearance properties
            of a clustered column chart.</p>
    </div>
    <table class="genericTable" style="font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" align="right">
                <img alt src="../../Image/ClusteredColumn.jpg"/></td>
            <td valign="top" align="left">
                <table class="genericTable">
                    <tr>
                        <td>
                            Category Axis Title:</td>
                        <td>
                            <asp:TextBox ID="CategoryAxisTitle" runat="server">Region</asp:TextBox></td>
                    </tr>
                    <tr>
                        <td>
                            Value Axis Title:</td>
                        <td>
                            <asp:TextBox ID="ValueAxisTitle" runat="server">In Thousands</asp:TextBox></td>
                    </tr>
                    <tr>
                        <td>
                            Value Axis MaxValue:</td>
                        <td>
                            <asp:DropDownList ID="ValueMaxValue" runat="server">
                                <asp:ListItem Value="0">80000</asp:ListItem>
                                <asp:ListItem Value="2">120000</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            Value Axis MinValue:</td>
                        <td>
                            <asp:DropDownList ID="ValueMinValue" runat="server">
                                <asp:ListItem Value="0">0</asp:ListItem>
                                <asp:ListItem Value="1">10000</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            Value Axis MajorUnit:</td>
                        <td>
                            <asp:DropDownList ID="ValueMajorUnit" runat="server">
                                <asp:ListItem Value="0">20000</asp:ListItem>
                                <asp:ListItem Value="1">40000</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            Value Axis MinorUnit</td>
                        <td>
                            <asp:DropDownList ID="ValueMinorUnit" runat="server">
                                <asp:ListItem Value="0">5000</asp:ListItem>
                                <asp:ListItem Value="0">10000</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            GapWidth:</td>
                        <td>
                            <asp:DropDownList ID="GapWidth" runat="server">
                                <asp:ListItem Value="1">100</asp:ListItem>
                                <asp:ListItem Value="2">200</asp:ListItem>
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
                        <td colspan="2">
                            <asp:Button ID="btnProcess" runat="server" Text="Create Report" OnClick="btnProcess_Click">
                            </asp:Button></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</asp:Content>
