<%@ Page Language="c#" Codebehind="stacked-bar.aspx.cs" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.StackedBar"
    MasterPageFile="~/tpl/Demo.Master" Title="Stacked Bar - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Stacked Bar - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
    <p>
        This Demo exhibits how to create a <b>Stacked Bar Chart</b> with 2-D and 3-D visual effects
        in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
            Aspose.Cells</a> for .NET.</p>
    <p>
        This type of chart shows the relationship of individual items to the whole. It is
        also available with a 3-D visual effect. Aspose.Cells is a powerful component, which
        supports all the standard and custom charts to help you display data in more meaningful
        ways. You may create many kinds of charts including Stacked Bar. The component can
        create the chart into the worksheet in a workbook using the simplest APIs with ease.
        The demo creates a workbook first and inputs the source data related chart into
        the first four columns (A, B, C and D) of the first worksheet named Data. The first
        column represents the category data (Region) where as the second, third and fourth
        columns represent the sales data representing values related to different products
        (Apple, Orange and Banana).
    </p>
    <p>
        The demo creates a Stacked Bar chart representing Fruit Sales By Region into the
        second worksheet named Chart based on the different product values related to different
        regions in the first worksheet. In the demo, you are provided a sample snapshot
        of the chart, a check box that represents whether you want to create the chart with
        3-D visual effect and a command button labeled Create Report to create and exercise
        the chart using your desired inputs. You can either open the resultant excel file
        into <b>MS Excel</b> or save directly to your disk to check the results.
    </p>
    <p>
        Click <b>Create Report</b> to see how demo can &nbsp;set the appearance properties
        of a stacked bar chart.</p>
        </div>
    <table class="genericTable" style="font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" align="right">
                <img alt="" src="../../Image/StackedBar.jpg" /></td>
            <td valign="top" align="left">
                <table class="genericTable">
                    <tr>
                        <td>
                            Show as 3D:</td>
                        <td>
                            <asp:CheckBox runat="server" ID="CheckBoxShow3D" /></td>
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
                            <asp:Button runat="server" ID="btnProcess" Text="Create Report" OnClick="btnProcess_Click">
                            </asp:Button></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</asp:Content>
