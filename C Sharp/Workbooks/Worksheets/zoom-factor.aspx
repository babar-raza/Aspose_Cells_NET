<%@ Page Language="c#" CodeBehind="zoom-factor.aspx.cs" AutoEventWireup="false" Inherits="Aspose.Cells.Demos.ZoomFactor"
    MasterPageFile="~/tpl/Demo.Master" Title="Zoom Factor - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_lft.jpg" width="19px" height="41px" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Zoom Factor - Aspose.Cells
                    </h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" src="/Common/images/heading_rt.jpg" width="19px" height="41px" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo exhibits how to set <b>Zoom</b> or <b>Scaling Factor</b> of the
            worksheets in a workbook using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                Aspose.Cells</a> for .NET.</p>
        <p>
            Aspose.Cells component provides the feature to set the <b>zoom</b> or <b>scaling factor</b>
            of the worksheets. The feature helps you to implement smaller or larger views to
            show the contents of your worksheet which may demand you at some occasions. The
            demo provides you a drop down list and a command button to implement the scaling
            factor. You may select a value from the drop down list which represents a percentage
            zoom factor value and click the button to generate the resultant report. It is to
            be noted that the percentage value you give should be <b>between 10 and 400</b>.
        </p>
        <p>
            Click <b>Create Report </b>tto see how example applies selected value a a <b>zoom factor</b>
            for the first worksheet of an excel document. You can either open the resulting
            excel file into <b>MS Excel</b> or save directly to your disk.
        </p>
        <table>
            <tr>
                <td>
                    <b>Zoom:&nbsp;</b>
                </td>
                <td>
                    <asp:DropDownList runat="server" ID="Zoom" Width="150">
                        <asp:ListItem Value="0">200</asp:ListItem>
                        <asp:ListItem Value="1">100</asp:ListItem>
                        <asp:ListItem Value="2">75</asp:ListItem>
                        <asp:ListItem Value="3">50</asp:ListItem>
                        <asp:ListItem Value="4">25</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <p>
                        <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                            <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                            <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
                        </asp:DropDownList>
                        <asp:Button runat="server" ID="Button1" Text="Create Report" />
                    </p>
                </td>
            </tr>
        </table>
    </div>
</asp:Content>
