<%@ Page Language="c#" Codebehind="cone-column.aspx.cs" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.ConeColumn"
    MasterPageFile="~/tpl/Demo.Master" Title="Cone Column - Aspose.Cells Demos" %>

<asp:Content runat="server" ContentPlaceHolderID="MainContent">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Cone Column - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo exhibits how to create a <b>Cone Column Chart</b> in a worksheet
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            The columns in these types of chart are represented by conical shapes. You may create
            it with stacked flavor too. Aspose.Cells is a powerful component, which supports
            all the standard and custom charts to help you display data in more meaningful ways.
            The demo creates a workbook first and inputs the chart source data into the first
            two columns (A and B) of the first worksheet named Cone Column. The first column
            represents the category data (Year spanned over 1996 - 2006) where as the second
            column represents the number of employees which denotes values in the chart.
        </p>
        <p>
            The demo creates a cone column chart representing number of employees into the worksheet
            based on the employee values related to different years. In the demo, you are provided
            a sample snapshot of the chart, a drop down list which represents the chart type
            (Cone and ConeStacked) and a command button labeled Create Report to create the
            chart based on your selection from the drop down list. The chart is created with
            a particular elevation and rotation angle with a specified depth and gap width in
            percentage b/w clusters. You can either open the resultant excel file into <b>MS Excel</b>
            or save directly to your disk.
        </p>
        <p>
            Click <b>Create Report</b> to see how demo can &nbsp;set the appearance properties
            of a cone column chart.</p>
    </div>
    <table class="genericTable" style="font-family: Arial; font-size: small;">
        <tr>
            <td align="right">
                <img alt="" src="../../Image/ConeColumn.jpg"/></td>
            <td valign="top" align="left">
                <table class="genericTable">
                    <tr>
                        <td>
                            Chart Type:
                        </td>
                        <td>
                            <asp:DropDownList runat="server" ID="ChartTypeList" Width="120">
                                <asp:ListItem Value="0">Cone</asp:ListItem>
                                <asp:ListItem Value="1">ConeStacked</asp:ListItem>
                            </asp:DropDownList></td>
                    </tr>
                    <tr>
                        <td>
                            Save Format:
                        </td>
                        <td>
                            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="120">
                                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" align="middle">
                            <asp:Button runat="server" ID="btnProcess" Text="Create Report" OnClick="btnProcess_Click" /></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</asp:Content>
