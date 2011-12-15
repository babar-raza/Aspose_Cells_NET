<%@ Page Language="c#" Codebehind="cylinderical-column-3d.aspx.cs" AutoEventWireup="false"
    Inherits="Aspose.Cells.Demos.CylindericalColumn3D" MasterPageFile="~/tpl/Demo.Master"
    Title="Cylinderical Column 3D - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Cylindrical Column 3D - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo exhibits how to create a Cylinder 3-D Column chart in a worksheet
            using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            The 3-D columns in these types of chart are represented by cylindrical shapes. Aspose.Cells
            is a powerful component, which supports all the standard and custom charts to help
            you display data in more meaningful ways. The demo creates a workbook first and
            inputs the chart source data into the first two columns (A and B) of the first worksheet
            named Cylindrical Column3D. The first column represents the category data (Year
            spanned over 1996 - 2006) where as the second column represents the number of employees
            which denotes values in the chart.
        </p>
        <p>
            The demo creates a cylinder 3-D column chart representing number of employees into
            the worksheet based on the employee values related to different years. In the demo,
            you are provided a sample snapshot of the chart and a command button labeled Create
            Report to create the chart. The chart is created with a particular elevation and
            rotation angle, a specified depth and gap width in percentage b/w clusters as well.
            You can either open the resultant excel file into MS Excel or save directly to your
            disk.
        </p>
        <p>
            Click <b>Create Report</b> to see how demo can &nbsp;set the appearance properties
            of a cylinder column 3D chart.</p>
    </div>
    <table class="genericTable" style="font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" align="right">
                <img alt="" src="../../Image/CylindericalColumn3D.jpg" /></td>
            <td valign="top" align="left">
                <table class="genericTable">
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
                        <td colspan="2">
                            <asp:Button runat="server" ID="btnProcess" Text="Create Report" OnClick="btnProcess_Click" /></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</asp:Content>
