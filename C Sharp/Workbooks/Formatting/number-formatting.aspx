<%@ Page Language="C#" MasterPageFile="~/tpl/Demo.Master" AutoEventWireup="true"
    CodeBehind="number-formatting.aspx.cs" Inherits="Workbooks_Formatting_NumberFormatting"
    Title="Number Formatting - Aspose.Cells Demos" %>

<asp:Content ID="Content" ContentPlaceHolderID="MainContent" runat="Server">
    <table width="90%" style="text-align: center" border="0" cellpadding="0" cellspacing="0">
        <tbody>
            <tr>
                <td valign="top" style="width: 19px">
                    <img alt="" height="41" src="/Common/images/heading_lft.jpg" width="19" />
                </td>
                <td class="demos-heading-bg" style="width: 100%">
                    <h2 class="demos-heading-bg">
                        Number Formatting - Aspose.Cells</h2>
                </td>
                <td valign="top" style="width: 19px">
                    <img alt="" height="41" src="/Common/images/heading_rt.jpg" width="19" />
                </td>
            </tr>
        </tbody>
    </table>
    <div style="text-align: left; font-family: Arial; font-size: small;">
        <p>
            This online demo exhibits how to apply <b>built-in Number formats</b> and <b>Custom
                Number formats</b> using <a href="http://www.aspose.com/categories/file-format-components/aspose.cells-for-.net-and-java/default.aspx">
                    Aspose.Cells</a> for .NET.</p>
        <p>
            You can use different Number formats to display data (<b>General format</b>, Numbers
            in <b>Decimal notations</b>, Numbers with <b>Currency symbols</b>, Numbers as a
            <b>Percentage of 100</b>, numbers in <b>Scientific format</b>, numbers in <b>DateTime
                values</b> and <b>Custom Number format</b>). The demo uses a template excel
            file with a number <b>1234.5</b> and then applies the values of different Number
            formats in column <b>B</b>. It also sets some Custom Number formats in Column <b>D</b>.
            For the complete list, Please see <a href="http://www.aspose.com/documentation/file-format-components/aspose.cells-for-.net-and-java/setting-display-formats-of-numbers-dates.html">
                Number Format List</a> . You can either open the resultant excel files into
            MS Excel or save directly to your disk to check the results.</p>
        <p>
            Please download the
            <asp:HyperLink ID="lnkFile" runat="server" NavigateUrl="~/designer/Workbooks/NumberFormatting.xls">NumberFormatting.xls</asp:HyperLink>
            used in this demo.
        </p>
        <p>
            <asp:DropDownList ID="ddlFileVersion" runat="server" Width="100">
                <asp:ListItem Selected="True" Value="XLS">XLS</asp:ListItem>
                <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
            </asp:DropDownList>
            <asp:Button ID="btnExecute" runat="server" Text="Process" OnClick="btnExecute_Click" />
        </p>
    </div>
</asp:Content>
