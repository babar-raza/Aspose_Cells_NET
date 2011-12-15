<%@ Page Language="c#" Codebehind="stacked-column.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.StackedColumn" MasterPageFile="~/tpl/Demo.Master"
    Title="StackedColumn - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Stacked Column - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo describes how to create a <b>Stacked Column Chart</b> with simple
            and 3-D visual effects in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            This type of chart shows the relationship of individual items to the whole, comparing
            the contribution of each value to a total across categories. Aspose.Cells is a powerful
            component, which supports all the standard and custom charts to help you display
            data in more meaningful ways. The demo creates a workbook first and inputs some
            chart related data into the first three columns (A, B and C) of the first worksheet
            named Data. The first column represents the category data (Year) where as the second
            and third column represent values for Product1 and Product2.
        </p>
        <p>
            The demo creates a stacked column chart representing product sales into the second
            worksheet named Chart based on the different product values related to different
            years (2004-2006) in the first worksheet. In the demo, you have been provided a
            sample snapshot of the chart, a check box that represents whether you want to create
            a 3-D stacked column chart and a command button labeled Create Report to create
            and exercise the chart using your desired inputs. You can either open the resultant
            excel file into <b>MS Excel</b> or save directly to your disk to check the results.
        </p>
        <p>
            Click <b>Create Report</b> to see how demo can set the appearance properties of
            a stacked column chart.</p>
    </div>
    <table class="genericTable" style="font-family: Arial; font-size: small;">
        <tr>
            <td align="right">
                <img alt="" src="../../Image/StackedColumn.jpg"></td>
            <td valign="top" align="left">
                <table class="genericTable">
                    <tr>
                        <td>
                            Show as 3D:</td>
                        <td>
                            <asp:CheckBox ID="checkBoxShow3D" runat="server"></asp:CheckBox></td>
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
