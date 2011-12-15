<%@ Page AutoEventWireup="false" Codebehind="exploded-pie.aspx.cs" Inherits="Aspose.Cells.Demos.ExplodedPie"
    Language="c#" MasterPageFile="~/tpl/Demo.Master" Title="ExplodedPie - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table class="componentDescriptionTxt" border="0" cellpadding="0" cellspacing="0"
        style="text-align: center; width: 100%; font-family: Arial; font-size: small;">
        <tr>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" /></td>
            <td class="demos-heading-bg" style="height: 41; width: 100%; font-family: Arial;
                font-size: large;">
                <h2 class="demos-heading-bg">
                    Exploded Pie - Aspose.Cells</h2>
            </td>
            <td valign="top" style="height: 41; width: 19">
                <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" /></td>
        </tr>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;" class="componentDescriptionTxt">
        <p>
            This online demo exhibits how to create an <b>Exploded Pie chart</b> with <b>2-D</b> and <b>3-D</b> visual
            effects in a worksheet using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            This type of chart displays the contribution of each value to a total while emphasizing
            individual values. It is also available with a 3-D visual effect. Aspose.Cells is
            a powerful component, which supports all the standard and custom charts to help
            you display data in more meaningful ways. The demo creates a workbook first and
            inputs the chart source data into the first two columns (A and B) of the first worksheet
            named Data. The first column represents the category data (Region) where as the
            second column represents the sales data representing values.
        </p>
        <p>
            The demo creates an exploded pie chart representing Sales By Region into the second
            worksheet named Chart based on the different sale values related to different regions
            in the first worksheet. In the demo, you are provided a sample snapshot of the chart,
            a check box that represents whether you want to create the chart with 3-D flavor
            and a command button labeled Create Report to create and exercise the chart using
            your desired inputs. You can either open the resultant excel file into <b>MS Excel</b>
            or save directly to your disk to check the results.
        </p>
        <p>
            Click <b>Create Report</b> to see how demo can set the appearance properties of
            a exploded pie chart.</p>
    </div>
    <table class="genericTable" style="font-family: Arial; font-size: small;">
        <tr>
            <td align="right">
                <img alt="" src="../../Image/ExplodedPie.jpg" /></td>
            <td valign="top" align="left">
                <table class="genericTable">
                    <tr>
                        <td>
                            Show as 3D:</td>
                        <td>
                            <asp:CheckBox runat="server" ID="CheckShow3D" /></td>
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
